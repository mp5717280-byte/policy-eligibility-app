# app.py
import streamlit as st
import pandas as pd
import re
from datetime import date

st.set_page_config(page_title="Policy Eligibility (Auto)", layout="centered")
st.title("Policy Eligibility Checker — Auto (Step 1 & Step 2)")

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

# ----------------------------
# Sidebar (only the essentials)
# ----------------------------
st.sidebar.header("Upload Files")
lic_file = st.sidebar.file_uploader("LIC 10 Year Plan (Excel)", type=["xlsx","xls"])
sheet2_file = st.sidebar.file_uploader("Excel Sheet 2 (Age rules)", type=["xlsx","xls"])

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
        st.error("Please upload both Excel files and enter the Policy Holder Name.")
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
                st.error("Could not detect UIN / Start / End in LIC file. Please ensure the table has UIN and From/To dates.")
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
                st.error("Could not read age rules from Excel Sheet 2. Ensure sheets are named like 'Policy 1', 'Policy 2', etc., with columns UIN / Minimum Entry Age / Maximum Entry Age.")
                st.stop()

            st.subheader("Step 2 — Age-Eligible Plans")
            final_df = step2_filter_by_age(step1, age_df, age_years)
            if final_df.empty:
                st.error("No plans eligible for this age based on Excel Sheet 2.")
            st.dataframe(final_df, use_container_width=True)

            st.download_button("Download Step-1 CSV", step1.to_csv(index=False).encode("utf-8"),
                               file_name="step1_active_plans.csv", mime="text/csv")
            st.download_button("Download Step-2 CSV (Final)", final_df.to_csv(index=False).encode("utf-8"),
                               file_name="step2_age_eligible_plans.csv", mime="text/csv")
        except Exception as e:
            st.exception(e)
else:
    st.info("Upload both Excel files, enter Name + DOB + Purchase Date, then click **Run Eligibility**.")
