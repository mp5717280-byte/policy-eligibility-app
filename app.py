# app.py
import streamlit as st
import pandas as pd
import re
from datetime import date
from io import BytesIO
from docx import Document  # needs python-docx in requirements.txt

st.set_page_config(page_title="Policy Eligibility (Auto + Step 3)", layout="centered")
st.title("Policy Eligibility Checker — Auto (Step 1, Step 2, Step 3)")

# =========================
# Helpers
# =========================
def calc_age_years(dob, purchase):
    """Age in years (decimal)."""
    return (purchase - dob).days / 365.25

def age_to_years(val):
    """Convert '90 days' / '3 months' / '18 years' / '18' → years (float)."""
    if val is None:
        return None
    if isinstance(val, float) and pd.isna(val):
        return None
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip().lower()
    m = re.search(r'([0-9]+(?:\.[0-9]+)?)', s)
    if not m:
        return None
    num = float(m.group(1))
    if "day" in s:
        return num / 365.0
    if "month" in s:
        return num / 12.0
    return num  # default years

def normalize_end_text_to_nat(series):
    """Treat words like 'Active'/'Present' as open-ended (NaT)."""
    def norm(v):
        if isinstance(v, str) and str(v).strip().lower() in {
            "active","ongoing","present","current","till date","tilldate","till-date"
        }:
            return pd.NaT
        return v
    return series.apply(norm)

def detect_header_row(raw_df, max_scan_rows=12):
    """Guess header row by scanning for tokens like UIN / From/To / Plan."""
    for i in range(min(max_scan_rows, len(raw_df))):
        vals = [str(v).strip().lower() for v in raw_df.iloc[i].values]
        if any("uin" in v for v in vals):
            return i
        if ("from" in vals and "to" in vals):
            return i
        if any(("plan" in v or "product" in v) for v in vals):
            return i
    return 0

def read_sheet_with_header(xls, sheet_name):
    raw = pd.read_excel(xls, sheet_name, header=None)
    h = detect_header_row(raw)
    df = pd.read_excel(xls, sheet_name, header=h)
    return df, h

def pick_best_lic_sheet(lic_xls):
    """Pick the sheet that best looks like the LIC plan table (has UIN and From/To)."""
    best_df, best_meta, best_score = None, None, -1
    for s in lic_xls.sheet_names:
        try:
            df, h = read_sheet_with_header(lic_xls, s)
        except Exception:
            continue
        if df is None or df.empty:
            continue
        df = df.loc[:, df.columns.notna()]
        cols_lower = [str(c).strip().lower() for c in df.columns]
        uin_col  = next((c for c,l in zip(df.columns, cols_lower) if "uin" in l), None)
        plan_col = next((c for c,l in zip(df.columns, cols_lower) if ("plan" in l or "product" in l) or ("name" in l and "column" not in l)), None)
        start_col = next((c for c,l in zip(df.columns, cols_lower) if l=="from" or "start" in l or "launch" in l or "effective" in l), None)
        end_col   = next((c for c,l in zip(df.columns, cols_lower) if l=="to"   or "end" in l or "withdraw" in l or "close" in l or "discontinu" in l), None)
        score = sum(x is not None for x in [uin_col, start_col, end_col]) + (1 if plan_col is not None else 0)
        if score > best_score or (score == best_score and len(df) > len(best_df or [])):
            best_df, best_meta, best_score = df, {
                "sheet": s, "header_row": h,
                "uin_col": uin_col, "plan_col": plan_col,
                "start_col": start_col, "end_col": end_col, "score": score
            }, score
    return best_df, best_meta

def step1_active_by_date(lic_df, meta, purchase_date):
    """Filter LIC plans active on the purchase date (no green highlight checks)."""
    if meta["uin_col"] is None or (meta["start_col"] is None and meta["end_col"] is None):
        return pd.DataFrame()

    start = pd.to_datetime(lic_df[meta["start_col"]], errors="coerce") if meta["start_col"] else pd.NaT
    end_raw = lic_df[meta["end_col"]] if meta["end_col"] else pd.Series([pd.NaT]*len(lic_df))
    end = pd.to_datetime(normalize_end_text_to_nat(end_raw), errors="coerce")
    pdate = pd.to_datetime(purchase_date)

    mask = pd.Series(True, index=lic_df.index)
    if meta["start_col"]: mask &= (start.isna() | (start <= pdate))
    if meta["end_col"]:   mask &= (end.isna()   | (end   >= pdate))

    cols = [c for c in [meta["uin_col"], meta["plan_col"], meta["start_col"], meta["end_col"]] if c]
    out = lic_df.loc[mask, cols].copy()
    out = out[out[meta["uin_col"]].notna() & (out[meta["uin_col"]].astype(str).str.strip()!="")]
    rename = {meta["uin_col"]: "UIN"}
    if meta["plan_col"]:  rename[meta["plan_col"]]  = "PlanName"
    if meta["start_col"]: rename[meta["start_col"]] = "StartDate"
    if meta["end_col"]:   rename[meta["end_col"]]   = "EndDate"
    return out.rename(columns=rename).drop_duplicates().reset_index(drop=True)

def build_age_table_auto(sheet2_xls):
    """Find all 'Policy...' sheets and extract UIN + min/max ages (to years)."""
    rows = []
    for s in sheet2_xls.sheet_names:
        if not s.lower().startswith("policy"):
            continue
        try:
            raw = pd.read_excel(sheet2_xls, s, header=None)
            h = detect_header_row(raw)
            df = pd.read_excel(sheet2_xls, s, header=h)
        except Exception:
            continue
        needed = ["UIN","Minimum Entry Age","Maximum Entry Age"]
        if not all(col in df.columns for col in needed):
            continue
        for _, r in df.iterrows():
            rows.append({
                "UIN": str(r["UIN"]).strip().upper(),
                "MinAgeYears": age_to_years(r["Minimum Entry Age"]),
                "MaxAgeYears": age_to_years(r["Maximum Entry Age"]),
                "Sheet": s
            })
    age_df = pd.DataFrame(rows)
    if age_df.empty:
        return age_df
    return age_df.dropna(subset=["UIN"]).drop_duplicates()

def step2_filter_by_age(step1_df, age_df, age_years):
    """Keep only Step-1 rows where age fits between Min/Max entry age."""
    if step1_df.empty or "UIN" not in step1_df.columns or age_df.empty:
        return pd.DataFrame()
    s1 = step1_df.copy()
    s1["UIN"] = s1["UIN"].astype(str).str.strip().str.upper()
    merged = pd.merge(s1, age_df, on="UIN", how="left")
    ok = merged[
        (merged["MinAgeYears"].notna()) &
        (merged["MaxAgeYears"].notna()) &
        (merged["MinAgeYears"] <= age_years) &
        (merged["MaxAgeYears"] >= age_years)
    ]
    return ok.drop_duplicates(subset=["UIN"]).reset_index(drop=True)

# ---- Step 3 helpers: UIN → Module map, DOCX parsing, dynamic fields ----
def find_uin_module_map(sheet2_xls):
    """
    Auto-find a sheet that has a UIN column and a Module column.
    Accepts column names like: Module, Module No, Module Name, Module ID, Input Module, Module #
    """
    module_aliases = ["module", "module no", "module name", "module id", "input module", "module #", "mod no", "mod"]
    for s in sheet2_xls.sheet_names:
        try:
            raw = pd.read_excel(sheet2_xls, s, header=None)
            h = detect_header_row(raw)
            df = pd.read_excel(sheet2_xls, s, header=h)
        except Exception:
            continue
        cols_lower = [str(c).strip().lower() for c in df.columns]
        if "uin" in cols_lower:
            module_idx = None
            for idx, l in enumerate(cols_lower):
                if any(alias in l for alias in module_aliases):
                    module_idx = idx
                    break
            if module_idx is not None:
                uin_col = df.columns[cols_lower.index("uin")]
                mod_col = df.columns[module_idx]
                m = df[[uin_col, mod_col]].dropna()
                m.columns = ["UIN", "Module"]
                m["UIN"] = m["UIN"].astype(str).str.strip().str.upper()
                m["Module"] = m["Module"].astype(str).str.strip()
                return m
    return pd.DataFrame(columns=["UIN","Module"])

def parse_client_input_fields_docx(doc_bytes: BytesIO):
    """
    Parse 'Client Input Fields.docx' into a dict: {'1': [fields...], '2': [...], ...}
    Heuristic: find lines like 'Module 1' then collect following lines until next module header.
    """
    doc = Document(doc_bytes)
    modules = {}
    current = None
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        m = re.match(r'^\s*module\s*(\d+)\b', text, flags=re.IGNORECASE)
        if m:
            current = m.group(1)
            if current not in modules:
                modules[current] = []
            continue
        if current:
            modules[current].append(text)
    # Clean up duplicates/short lines
    for k in list(modules.keys()):
        cleaned, seen = [], set()
        for line in modules[k]:
            norm = line.strip()
            if len(norm) < 2:
                continue
            if norm in seen:
                continue
            seen.add(norm)
            cleaned.append(norm)
        modules[k] = cleaned
    return modules

def render_dynamic_inputs(fields_list):
    """
    Render inputs per field label.
    - If contains 'date' -> date_input
    - If contains 'select'/'dropdown' -> text_input with hint
    - Else -> text_input
    """
    out = {}
    for label in fields_list:
        low = label.lower()
        key = f"fld::{label}"
        if "date" in low:
            out[label] = st.date_input(label, key=key)
        elif "select" in low or "dropdown" in low or "choose" in low:
            out[label] = st.text_input(label + " (type your choice)", key=key)
        else:
            out[label] = st.text_input(label, key=key)
    return out

# =========================
# Sidebar (inputs)
# =========================
st.sidebar.header("Upload Files")
lic_file = st.sidebar.file_uploader("LIC 10 Year Plan (Excel)", type=["xlsx","xls"])
sheet2_file = st.sidebar.file_uploader("Excel Sheet 2 (Age rules + UIN→Module map)", type=["xlsx","xls"])
client_docx = st.sidebar.file_uploader("Client Input Fields (DOCX)", type=["docx"])

st.sidebar.header("Client Inputs")
name = st.sidebar.text_input("Policy Holder Name", "")
dob = st.sidebar.date_input("Date of Birth", value=date(2000,1,1))
purchase = st.sidebar.date_input("Policy Purchase Date", value=date(2020,3,20))
run = st.sidebar.button("Run Eligibility")

# =========================
# Main
# =========================
if run:
    if not lic_file or not sheet2_file or not name:
        st.error("Please upload LIC Excel, Sheet 2 Excel, and enter the Policy Holder Name.")
    else:
        age_years = calc_age_years(dob, purchase)
        st.subheader("Inputs")
        st.write(pd.DataFrame([{
            "Policy Holder": name,
            "DOB": dob.isoformat(),
            "Purchase Date": purchase.isoformat(),
            "Age at Purchase (years)": round(age_years, 4)
        }]))

        # ----- Step 1: Auto detect sheet/headers and filter by date -----
        try:
            lic_xls = pd.ExcelFile(lic_file)
            lic_df, meta = pick_best_lic_sheet(lic_xls)
            if lic_df is None or meta is None or meta.get("score", 0) < 2:
                st.error("Could not detect UIN / Start / End in LIC file. Ensure the table has UIN and From/To dates.")
                st.stop()

            step1 = step1_active_by_date(lic_df, meta, purchase)
            st.subheader("Step 1 — Active Plans on Purchase Date")
            st.info(f"Detected sheet: {meta['sheet']} (header row ~ {meta['header_row']+1})")
            st.write({"uin_col": meta["uin_col"], "plan_col": meta["plan_col"],
                      "start_col": meta["start_col"], "end_col": meta["end_col"], "rows": len(step1)})
            if step1.empty:
                st.warning("No plans found for that date. Check your LIC file's From/To dates.")
            st.dataframe(step1, use_container_width=True)
        except Exception as e:
            st.exception(e)
            st.stop()

        # ----- Step 2: Build age table and filter by age -----
        try:
            s2_xls = pd.ExcelFile(sheet2_file)
            age_df = build_age_table_auto(s2_xls)
            if age_df.empty:
                st.error("Could not read age rules from Excel Sheet 2. Ensure 'Policy ...' sheets have UIN / Minimum / Maximum Entry Age.")
                st.stop()

            st.subheader("Step 2 — Age-Eligible Plans")
            final_df = step2_filter_by_age(step1, age_df, age_years)
            if final_df.empty:
                st.error("No plans eligible for this age based on Excel Sheet 2.")
            st.dataframe(final_df, use_container_width=True)

            # Select a plan (UIN) for Step 3
            selected_uin = None
            if not final_df.empty:
                plan_names = final_df.get("PlanName", pd.Series([""]*len(final_df))).fillna("")
                options = (final_df["UIN"].astype(str).str.strip().str.upper() + " — " + plan_names.astype(str)).tolist()
                choice = st.selectbox("Select a plan for Step 3", options) if options else None
                if choice:
                    selected_uin = choice.split(" — ")[0].strip()

            # Download buttons
            st.download_button("Download Step-1 CSV", step1.to_csv(index=False).encode("utf-8"),
                               file_name="step1_active_plans.csv", mime="text/csv")
            st.download_button("Download Step-2 CSV (Final)", final_df.to_csv(index=False).encode("utf-8"),
                               file_name="step2_age_eligible_plans.csv", mime="text/csv")
        except Exception as e:
            st.exception(e)
            st.stop()

        # ----- Step 3: UIN → Module → Fields -----
        if selected_uin:
            st.subheader("Step 3 — Module & Client Input Fields")
            # 3a) Try auto UIN→Module mapping
            uin_module_map = find_uin_module_map(s2_xls)

            # 3b) If auto fails, manual picker (sheet + columns)
            if uin_module_map.empty:
                st.warning("Auto-detect could not find a UIN → Module table. Pick it manually below.")
                s2_sheet = st.selectbox("Choose the Sheet in Excel Sheet 2 that has the UIN→Module mapping",
                                        s2_xls.sheet_names, key="s2map_sheet")
                try:
                    raw = pd.read_excel(s2_xls, s2_sheet, header=None)
                    h = detect_header_row(raw)
                    df_map = pd.read_excel(s2_xls, s2_sheet, header=h)
                    st.caption("Columns in this sheet:")
                    st.write(list(df_map.columns))

                    uin_col_pick = st.selectbox("Which column is UIN?", list(df_map.columns), key="s2map_uin")
                    module_col_pick = st.selectbox("Which column is Module?", list(df_map.columns), key="s2map_mod")

                    if st.button("Use this mapping"):
                        tmp = df_map[[uin_col_pick, module_col_pick]].dropna().copy()
                        tmp.columns = ["UIN", "Module"]
                        tmp["UIN"] = tmp["UIN"].astype(str).str.strip().str.upper()
                        tmp["Module"] = tmp["Module"].astype(str).str.strip()
                        uin_module_map = tmp
                        st.success("Mapping set successfully.")
                except Exception as e:
                    st.exception(e)

            if uin_module_map.empty:
                st.stop()

            # 3c) Look up selected UIN's module
            uin_clean = selected_uin.strip().upper()
            row = uin_module_map[uin_module_map["UIN"] == uin_clean]
            if row.empty:
                st.error(f"No Module mapping found for UIN: {selected_uin}")
                st.stop()
            module_raw = str(row.iloc[0]["Module"]).strip()
            mm = re.search(r'(\d+)', module_raw)
            module_id = mm.group(1) if mm else module_raw
            st.success(f"Detected Module: {module_id} (from Excel Sheet 2)")

            # 3d) Parse DOCX and render fields
            if not client_docx:
                st.warning("Upload the 'Client Input Fields.docx' in the sidebar to show Module-specific fields.")
            else:
                fields_by_module = parse_client_input_fields_docx(client_docx)
                fields_list = fields_by_module.get(module_id, [])
                if not fields_list:
                    st.warning(f"No fields found in DOCX for Module {module_id}.")
                else:
                    st.write(f"Fields required for Module {module_id}:")
                    user_inputs = render_dynamic_inputs(fields_list)
                    st.info("These values can be used in the next steps of your policy review engine.")
                    st.dataframe(pd.DataFrame([user_inputs]), use_container_width=True)

else:
    st.info("Upload LIC & Sheet 2 Excel files (and DOCX optionally). Enter Name + DOB + Purchase Date, then click **Run Eligibility**.")
