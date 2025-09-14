# app.py
import streamlit as st
import pandas as pd
import re
from datetime import date
from io import BytesIO

# NEW: for reading the Word doc with module fields
from docx import Document

st.set_page_config(page_title="Policy Eligibility (Auto + Step 3)", layout="centered")
st.title("Policy Eligibility Checker — Auto (Step 1, Step 2, Step 3)")

# ----------------------------
# Helpers
# ----------------------------
def calc_age_years(dob, purchase):
    return (purchase - dob).days / 365.25

def age_to_years(val):
    if val is None or (isinstance(val, float) and pd.isna(val)):
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
    return num  # default: years

def normalize_end_text_to_nat(series):
    def norm(v):
        if isinstance(v, str) and str(v).strip().lower() in {
            "active","ongoing","present","current","till date","tilldate","till-date"
        }:
            return pd.NaT
        return v
    return series.apply(norm)

def detect_header_row(raw_df, max_scan_rows=12):
    # Look for a row that contains header-like tokens
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
    """Pick the sheet that best looks like the LIC plan table (UIN/From/To)."""
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
    """Filter LIC plans active on purchase_date (no green)."""
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

# ---------- NEW (STEP 3): UIN → Module → Fields ----------
def find_uin_module_map(sheet2_xls):
    """
    Look for a sheet that contains a UIN-to-Module map.
    Heuristics: any sheet that has a 'UIN' column AND a column whose header contains 'module'.
    """
    for s in sheet2_xls.sheet_names:
        try:
            raw = pd.read_excel(sheet2_xls, s, header=None)
            h = detect_header_row(raw)
            df = pd.read_excel(sheet2_xls, s, header=h)
        except Exception:
            continue
        cols_lower = [str(c).strip().lower() for c in df.columns]
        if "uin" in cols_lower and any("module" in l for l in cols_lower):
            uin_col = df.columns[cols_lower.index("uin")]
            mod_col = next(df.columns[i] for i,l in enumerate(cols_lower) if "module" in l)
            m = df[[uin_col, mod_col]].dropna()
            m.columns = ["UIN", "Module"]
            # normalize
            m["UIN"] = m["UIN"].astype(str).str.strip().str.upper()
            # module as string like "Module 1" or just number
            m["Module"] = m["Module"].astype(str).str.strip()
            return m
    return pd.DataFrame(columns=["UIN","Module"])

def parse_client_input_fields_docx(doc_bytes: BytesIO):
    """
    Parse 'Client Input Fields.docx' into a dict: {'1': [fields...], '2': [...], ...}
    Heuristic parser: looks for 'Module 1', 'Module 2', ..., captures following paragraph lines
    until next 'Module N'.
    """
    doc = Document(doc_bytes)
    modules = {}
    current = None
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
        # detect a module header like "Module 1" or "Module 1 – ..." (robust)
        m = re.match(r'^\s*module\s*(\d+)\b', text, flags=re.IGNORECASE)
        if m:
            current = m.group(1)
            if current not in modules:
                modules[current] = []
            continue
        if current:
            # treat bullet points / numbered lines / plain lines as fields
            modules[current].append(text)
    # Cleanup: drop duplicate lines and very short separators
    for k in list(modules.keys()):
        cleaned = []
        seen = set()
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
    Render input widgets based on field label heuristics.
    - If contains 'date' -> date_input
    - If contains 'select' or 'dropdown' -> text_input placeholder (choices unknown)
    - Else -> text_input
    Returns dict of {label: value}
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

# ----------------------------
# Sidebar (only the essentials)
# ----------------------------
st.sidebar.header("Upload Files")
lic_file = st.sidebar.file_uploader("LIC 10 Year Plan (Excel)", type=["xlsx","xls"])
sheet2_file = st.sidebar.file_uploader("Excel Sheet 2 (Age rules + UIN→Module map)", type=["xlsx","xls"])
# NEW: Word doc with module-specific client inputs
client_docx = st.sidebar.file_uploader("Client Input Fields (DOCX)", type=["docx"])

st.sidebar.header("Client Inputs")
name = st.sidebar.text_input("Policy Holder Name", "")
dob = st.sidebar.date_input("Date of Birth", value=date(2000,1,1))
purchase = st.sidebar.date_input("Policy Purchase Date", value=date(2020,3,20))
run = st.sidebar.button("Run Eligibility")

# ----------------------------
# Main
# ----------------------------
if run:
    if not lic_file or not sheet2_file or not name:
        st.error("Please upload LIC Excel, Sheet 2 Excel, enter the Policy Holder Name.")
    else:
        age_years = calc_age_years(dob, purchase)
        st.subheader("Inputs")
        st.write(pd.DataFrame([{
            "Policy Holder": name,
            "DOB": dob.isoformat(),
            "Purchase Date": purchase.isoformat(),
            "Age at Purchase (years)": round(age_years, 4)
        }]))

        # Step 1 — Auto-detect sheet + header + columns
        try:
            lic_xls = pd.ExcelFile(lic_file)
            lic_df, meta = pick_best_lic_sheet(lic_xls)
            if lic_df is None or meta is None or meta.get("score", 0) < 2:
                st.error("Could not detect UIN / Start / End in LIC file. Ensure the table has UIN and From/To dates.")
                st.stop()

            step1 = step1_active_by_date(lic_df, meta, purchase)
            st.subheader("Step 1 — Active Plans on Purchase Date")
            st.info(f"Detected sheet: {meta['sheet']} (header row ~ {meta['header_row']+1})")
            st.write({"uin_col": meta["uin_col"], "plan_col": meta["plan_col"], "start_col": meta["start_col"], "end_col": meta["end_col"], "rows": len(step1)})
            if step1.empty:
                st.warning("No plans found for that date. Check your LIC file's From/To dates.")
            st.dataframe(step1, use_container_width=True)
        except Exception as e:
            st.exception(e)
            st.stop()

        # Step 2 — Auto-build age table & apply
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

            # Allow the user to select one plan for Step 3
            selected_uin = None
            if not final_df.empty:
                options = (final_df["UIN"] + " — " + final_df.get("PlanName", pd.Series([""]*len(final_df))).fillna("")).tolist()
                choice = st.selectbox("Select a plan for Step 3", options) if options else None
                if choice:
                    selected_uin = choice.split(" — ")[0].strip()

            st.download_button("Download Step-1 CSV", step1.to_csv(index=False).encode("utf-8"),
                               file_name="step1_active_plans.csv", mime="text/csv")
            st.download_button("Download Step-2 CSV (Final)", final_df.to_csv(index=False).encode("utf-8"),
                               file_name="step2_age_eligible_plans.csv", mime="text/csv")
        except Exception as e:
            st.exception(e)
            st.stop()

        # ---------- STEP 3: Module detection + show module fields ----------
        if selected_uin:
            st.subheader("Step 3 — Module & Client Input Fields")
            try:
                # 1) Build UIN → Module mapping from Sheet 2
                uin_module_map = find_uin_module_map(s2_xls)
                if uin_module_map.empty:
                    st.error("Could not find UIN→Module mapping in Excel Sheet 2 (look for a sheet with UIN and Module columns).")
                else:
                    # Normalize selected
                    uin_clean = selected_uin.strip().upper()
                    row = uin_module_map[uin_module_map["UIN"] == uin_clean]
                    if row.empty:
                        st.error(f"No Module mapping found for UIN: {selected_uin}")
                    else:
                        module_raw = str(row.iloc[0]["Module"]).strip()
                        # Extract numeric module id (supports 'Module 1', '1', etc.)
                        mm = re.search(r'(\d+)', module_raw)
                        module_id = mm.group(1) if mm else module_raw
                        st.success(f"Detected Module: {module_id} (from Excel Sheet 2)")

                        # 2) Parse the Client Input Fields DOCX
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
                                # (Optional) Show collected values as a small table
                                show_df = pd.DataFrame([user_inputs])
                                st.dataframe(show_df, use_container_width=True)
            except Exception as e:
                st.exception(e)

else:
    st.info("Upload LIC & Sheet 2 Excel files, plus (optional) Client Input Fields.docx. Enter Name + DOB + Purchase Date, then click **Run Eligibility**.")
