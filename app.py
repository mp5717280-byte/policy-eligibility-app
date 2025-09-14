# app.py
import streamlit as st
import pandas as pd
import numpy as np
import re
from datetime import date

st.set_page_config(page_title="Policy Eligibility (Step 1 & 2)", layout="centered")
st.title("Policy Eligibility Checker — Step 1 & Step 2 (Clean Restart)")

# ----------------------------
# Helpers
# ----------------------------
def calc_age_years(dob: date, purchase: date) -> float:
    """Age in years (decimal)."""
    return (purchase - dob).days / 365.25

def age_to_years(val):
    """Convert '90 days' / '3 months' / '18 years' / '18' → years (float)."""
    if pd.isna(val):
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
    # default assume years
    return num

def normalize_end_text_to_nat(series: pd.Series) -> pd.Series:
    """Treat text like 'Active', 'Present' as blank end-date (still active)."""
    return series.apply(
        lambda v: pd.NaT if isinstance(v, str) and str(v).strip().lower()
        in {"active","ongoing","present","current","till date","tilldate","till-date"} else v
    )

def build_step1_shortlist(lic_df: pd.DataFrame, uin_col, plan_col, start_col, end_col, purchase_date):
    """Filter LIC plans active on purchase_date. (No green highlight to keep simple & robust)"""
    # Parse dates
    start = pd.to_datetime(lic_df[start_col], errors="coerce") if start_col else pd.NaT
    end_raw = lic_df[end_col] if end_col else pd.Series([pd.NaT] * len(lic_df))
    end = pd.to_datetime(normalize_end_text_to_nat(end_raw), errors="coerce")
    pdate = pd.to_datetime(purchase_date)

    mask = pd.Series(True, index=lic_df.index)
    if start_col: mask &= (start.isna() | (start <= pdate))
    if end_col:   mask &= (end.isna()   | (end   >= pdate))

    cols = [c for c in [uin_col, plan_col, start_col, end_col] if c]
    out = lic_df.loc[mask, cols].copy()

    # Clean & rename
    out = out[out[uin_col].notna() & (out[uin_col].astype(str).str.strip() != "")]
    rename = {uin_col: "UIN"}
    if plan_col:  rename[plan_col]  = "PlanName"
    if start_col: rename[start_col] = "StartDate"
    if end_col:   rename[end_col]   = "EndDate"
    out = out.rename(columns=rename).drop_duplicates().reset_index(drop=True)
    return out

def build_age_table(sheet2_xls: pd.ExcelFile, header_row_index: int) -> pd.DataFrame:
    """
    Combine all 'Policy ...' sheets into one UIN/min/max table.
    header_row_index is 0-based (use 1 for 2nd row in Excel).
    """
    rows = []
    for s in sheet2_xls.sheet_names:
        if not s.lower().startswith("policy"):
            continue
        try:
            df = pd.read_excel(sheet2_xls, s, header=header_row_index)
        except Exception:
            continue
        needed = ["UIN", "Minimum Entry Age", "Maximum Entry Age"]
        if not all(c in df.columns for c in needed):
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
    age_df = age_df.dropna(subset=["UIN"]).drop_duplicates()
    return age_df

def apply_age_filter(step1_df: pd.DataFrame, age_df: pd.DataFrame, age_years: float) -> pd.DataFrame:
    """Filter Step-1 list by age eligibility."""
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
# Sidebar Inputs (simple & manual)
# ----------------------------
st.sidebar.header("Upload Files")
lic_file = st.sidebar.file_uploader("LIC 10 Year Plan (Excel)", type=["xlsx","xls"])
sheet2_file = st.sidebar.file_uploader("Excel Sheet 2 (Age rules)", type=["xlsx","xls"])

st.sidebar.header("Client Inputs")
name = st.sidebar.text_input("Policy Holder Name", "")
dob = st.sidebar.date_input("Date of Birth", value=date(2000, 1, 1))
purchase = st.sidebar.date_input("Policy Purchase Date", value=date(2020, 3, 20))

# Manual mapping always (to avoid auto-detect confusion)
st.sidebar.header("LIC Sheet Settings")
lic_sheet_name = st.sidebar.text_input("LIC sheet name (exact as in Excel)", value="")
lic_header_row_1based = st.sidebar.number_input("LIC header row (1-based)", min_value=1, value=2, step=1,
                                                help="If your headers are on the 2nd Excel row, type 2 here.")

st.sidebar.header("Sheet 2 Settings")
sheet2_header_row_1based = st.sidebar.number_input("Sheet 2 header row (1-based)", min_value=1, value=2, step=1,
                                                   help="Commonly 2 for your file.")

run = st.sidebar.button("Run Eligibility")

# ----------------------------
# Main Flow
# ----------------------------
if run:
    if not lic_file or not sheet2_file or not name or not lic_sheet_name.strip():
        st.error("Please upload both files, enter Policy Holder Name, and the LIC sheet name.")
    else:
        # Age
        age_years = calc_age_years(dob, purchase)
        st.subheader("Inputs")
        st.write(pd.DataFrame([{
            "Policy Holder": name,
            "DOB": dob.isoformat(),
            "Purchase Date": purchase.isoformat(),
            "Age at Purchase (years)": round(age_years, 4)
        }]))

        # ------------ Read & map LIC sheet (manual, reliable) ------------
        try:
            lic_xls = pd.ExcelFile(lic_file)
            if lic_sheet_name not in lic_xls.sheet_names:
                st.error(f"LIC sheet '{lic_sheet_name}' not found. Sheets available: {lic_xls.sheet_names}")
                st.stop()

            # Preview (no header) to help user confirm header row visually
            lic_preview = pd.read_excel(lic_xls, lic_sheet_name, header=None)
            st.caption("Preview first 15 rows (no header) to help you choose correct header row:")
            st.dataframe(lic_preview.head(15), use_container_width=True)

            lic_header_idx = int(lic_header_row_1based) - 1
            lic_df = pd.read_excel(lic_xls, lic_sheet_name, header=lic_header_idx)
            st.success(f"Loaded LIC sheet '{lic_sheet_name}' with header row = {lic_header_row_1based}")
            st.write("Detected columns:", list(lic_df.columns))

            # Column mapping dropdowns
            uin_col   = st.selectbox("Column for UIN", list(lic_df.columns))
            plan_col  = st.selectbox("Column for Plan Name", list(lic_df.columns))
            start_col = st.selectbox("Column for Start/From", list(lic_df.columns))
            end_col   = st.selectbox("Column for End/To", list(lic_df.columns))

            # Step 1
            step1 = build_step1_shortlist(lic_df, uin_col, plan_col, start_col, end_col, purchase)
            st.subheader("Step 1 — Active Plans on Purchase Date")
            if step1.empty:
                st.warning("No plans found in Step-1. Check header row & column mapping.")
            st.dataframe(step1, use_container_width=True)

        except Exception as e:
            st.exception(e)
            st.stop()

        # ------------ Read Sheet 2 & apply age rules ------------
        try:
            s2_header_idx = int(sheet2_header_row_1based) - 1
            sheet2_xls = pd.ExcelFile(sheet2_file)
            age_df = build_age_table(sheet2_xls, header_row_index=s2_header_idx)
            if age_df.empty:
                st.error("Could not build Age table from Sheet 2. Check header row & that sheets are named like 'Policy 1', 'Policy 2', etc.")
                st.stop()

            # Step 2
            st.subheader("Step 2 — Age-Eligible Plans")
            final_df = apply_age_filter(step1, age_df, age_years)
            if final_df.empty:
                st.error("No plans eligible for this age based on Excel Sheet 2.")
            st.dataframe(final_df, use_container_width=True)

            # Downloads
            st.download_button("Download Step-1 CSV", step1.to_csv(index=False).encode("utf-8"),
                               file_name="step1_active_plans.csv", mime="text/csv")
            st.download_button("Download Step-2 CSV (Final)", final_df.to_csv(index=False).encode("utf-8"),
                               file_name="step2_age_eligible_plans.csv", mime="text/csv")

        except Exception as e:
            st.exception(e)

else:
    st.info("Upload both Excel files, enter inputs, type the LIC sheet name, set header rows, then click **Run Eligibility**.")
