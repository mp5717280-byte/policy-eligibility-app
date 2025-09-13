

st.set_page_config(page_title="LIC Plan Eligibility (Step 1 & 2)", layout="centered")

st.title("LIC Plan Eligibility — Step 1 & 2")
st.caption("Upload your Excel files once, enter client details, and get the final eligible plans.")

# -------- Helpers --------
def calc_age_years(dob: date, purchase: date) -> float:
    rd = relativedelta(purchase, dob)
    return rd.years + rd.months/12 + rd.days/365.0

def normalize_date_any(x):
    if pd.isna(x):
        return pd.NaT
    try:
        return pd.to_datetime(x, errors="coerce")
    except Exception:
        return pd.NaT

def age_to_years(val):
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
    return num  # default years

def is_greenish(color: Color) -> bool:
    if not color or not isinstance(color.rgb, str) or not color.rgb:
        return False
    rgb = color.rgb
    if len(rgb) == 8:  # ARGB
        hex_rgb = rgb[2:]
    elif len(rgb) == 6:
        hex_rgb = rgb
    else:
        return False
    try:
        r = int(hex_rgb[0:2], 16); g = int(hex_rgb[2:4], 16); b = int(hex_rgb[4:6], 16)
        return (g >= 140) and (g > r + 20) and (g > b + 20)
    except Exception:
        return False

def detect_green_rows_xlsx(file, sheet_name: str, max_cols: int = 60):
    # file is a BytesIO uploaded file; openpyxl can load it directly
    wb = load_workbook(file, data_only=True)
    ws = wb[sheet_name]
    greens = set()
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row, min_col=1, max_col=min(max_cols, ws.max_column)):
        rindex = row[0].row
        if any(is_greenish(c.fill.fgColor) for c in row if c.fill):
            greens.add(rindex)
    return greens

def autodetect_plan_sheet_and_columns(df_dict):
    # df_dict: {sheet_name: dataframe_with_headers}
    best = None
    best_meta = None
    for s, df in df_dict.items():
        if df is None or df.empty: 
            continue
        df = df.loc[:, df.columns.notna()]
        cols_lower = [str(c).strip().lower() for c in df.columns]
        uin_col  = next((c for c,l in zip(df.columns, cols_lower) if "uin" in l), None)
        plan_col = next((c for c,l in zip(df.columns, cols_lower) if ("plan" in l or "product" in l) or ("name" in l and "column" not in l)), None)
        start_col = next((c for c,l in zip(df.columns, cols_lower) if any(t in l for t in ["start","launch","effective","from"])), None)
        end_col   = next((c for c,l in zip(df.columns, cols_lower) if any(t in l for t in ["end","withdraw","close","to","discontinu"])), None)
        score = sum(x is not None for x in [uin_col, start_col, end_col]) + (1 if plan_col is not None else 0)
        if best is None or score > best_meta["score"] or (score == best_meta["score"] and len(df) > len(best)):
            best = df
            best_meta = {"sheet": s, "uin_col": uin_col, "plan_col": plan_col, "start_col": start_col, "end_col": end_col, "score": score}
    return best, best_meta

def build_date_shortlist(lic_file, lic_df_dict, name, dob, purchase_date, use_green=True):
    # autodetect columns
    df, meta = autodetect_plan_sheet_and_columns(lic_df_dict)
    if df is None or meta["score"] < 2:
        st.error("Could not detect UIN/Start/End columns in the LIC file. Check headers.")
        return pd.DataFrame(), meta

    # parse dates
    start = pd.to_datetime(df[meta["start_col"]], errors="coerce") if meta["start_col"] else pd.NaT
    end_raw = df[meta["end_col"]] if meta["end_col"] else pd.Series([pd.NaT]*len(df))
    end_raw = end_raw.apply(lambda v: pd.NaT if isinstance(v, str) and str(v).strip().lower() in {"active","ongoing","present","current","till date","tilldate","till-date"} else v)
    end = pd.to_datetime(end_raw, errors="coerce")
    pdate = pd.to_datetime(purchase_date)

    # green highlight detection (optional)
    df = df.reset_index(drop=True)
    if use_green:
        greens = detect_green_rows_xlsx(lic_file, meta["sheet"])
        # Worksheet row = df index + header offset (assume header in first non-empty row)
        # We cannot know exact header offset here; a pragmatic approximation is +2, works for most tables with header at row 1.
        df["__ws_row__"] = df.index + 2
        df["__is_green__"] = df["__ws_row__"].isin(greens)
        mask = df["__is_green__"]
    else:
        mask = pd.Series(True, index=df.index)

    if meta["start_col"]: mask &= (start.isna() | (start <= pdate))
    if meta["end_col"]:   mask &= (end.isna()   | (end   >= pdate))

    out_cols = [c for c in [meta["uin_col"], meta["plan_col"], meta["start_col"], meta["end_col"]] if c]
    short = df.loc[mask, out_cols].copy()
    if meta["uin_col"]:
        short = short[short[meta["uin_col"]].notna() & (short[meta["uin_col"]].astype(str).str.strip()!="")]
    rename_map = {}
    if meta["uin_col"]:   rename_map[meta["uin_col"]]   = "UIN"
    if meta["plan_col"]:  rename_map[meta["plan_col"]]  = "PlanName"
    if meta["start_col"]: rename_map[meta["start_col"]] = "StartDate"
    if meta["end_col"] :  rename_map[meta["end_col"]]   = "EndDate"
    short = short.rename(columns=rename_map).drop_duplicates().reset_index(drop=True)
    return short, meta

def age_filter(short_df, sheet2_dict, age_years: float):
    # Build a combined age table from all 'Policy X' sheets
    rows = []
    for s, df in sheet2_dict.items():
        if not s.lower().startswith("policy"):
            continue
        if df is None or df.empty:
            continue
        if not all(c in df.columns for c in ["UIN","Minimum Entry Age","Maximum Entry Age"]):
            continue
        for _, r in df.iterrows():
            rows.append({
                "UIN": str(r["UIN"]).strip().upper(),
                "MinAgeYears": age_to_years(r["Minimum Entry Age"]),
                "MaxAgeYears": age_to_years(r["Maximum Entry Age"]),
                "Sheet": s
            })
    age_df = pd.DataFrame(rows).dropna(subset=["UIN"]).drop_duplicates()

    # Normalize shortlist UIN
    short_df = short_df.copy()
    if "UIN" not in short_df.columns or short_df.empty:
        return pd.DataFrame()
    short_df["UIN"] = short_df["UIN"].astype(str).str.strip().str.upper()

    merged = pd.merge(short_df, age_df, on="UIN", how="left")
    eligible = merged[(merged["MinAgeYears"].notna()) & (merged["MaxAgeYears"].notna()) & (merged["MinAgeYears"] <= age_years) & (merged["MaxAgeYears"] >= age_years)]
    eligible = eligible.drop_duplicates(subset=["UIN"]).reset_index(drop=True)
    return eligible

# -------- Sidebar Inputs --------
st.sidebar.header("Upload Files")
lic_file = st.sidebar.file_uploader("LIC 10 Year Plan for Developer (Excel)", type=["xlsx"])
sheet2_file = st.sidebar.file_uploader("Excel Sheet 2 (Age Eligibility)", type=["xlsx"])

st.sidebar.header("Client Inputs")
name = st.sidebar.text_input("Policy Holder Name", "")
dob = st.sidebar.date_input("Date of Birth", value=date(2000,1,1))
purchase = st.sidebar.date_input("Policy Purchase Date", value=date(2020,3,20))
use_green = st.sidebar.checkbox("Use GREEN highlight filter", value=True)

run = st.sidebar.button("Run Eligibility")


# -------- Main --------
if run:
    if not lic_file or not sheet2_file or not name:
        st.warning("Please upload both Excel files and enter the Policy Holder Name.")
    else:
        # Read all sheets as DataFrames (header=0 best-effort)
        lic_xls = pd.ExcelFile(lic_file)
        lic_df_dict = {s: pd.read_excel(lic_xls, s) for s in lic_xls.sheet_names}
        sheet2_xls = pd.ExcelFile(sheet2_file)
        sheet2_dict = {s: pd.read_excel(sheet2_xls, s, header=1) for s in sheet2_xls.sheet_names}

        age_years = calc_age_years(dob, purchase)
        st.subheader("Inputs")
        st.write(pd.DataFrame([{
            "Policy Holder": name,
            "DOB": dob.isoformat(),
            "Purchase Date": purchase.isoformat(),
            "Age at Purchase (years)": round(age_years, 4)
        }]))

        st.subheader("Step 1 — Active Plans by Date")
        short, meta = build_date_shortlist(lic_file, lic_df_dict, name, dob, purchase, use_green=use_green)
        if short.empty:
            st.warning("No plans found for that date with current settings. Try disabling GREEN filter or check sheet headers.")
        st.dataframe(short)

        st.subheader("Step 2 — Age-Eligible Plans")
        final_df = age_filter(short, sheet2_dict, age_years)
        if final_df.empty:
            st.error("No plans eligible for this age based on Excel Sheet 2.")
        st.dataframe(final_df)

        # Download buttons
        st.download_button("Download Step-1 CSV", short.to_csv(index=False).encode("utf-8"), file_name="step1_active_plans.csv", mime="text/csv")
        st.download_button("Download Step-2 CSV (Final)", final_df.to_csv(index=False).encode("utf-8"), file_name="step2_age_eligible_plans.csv", mime="text/csv")

        st.success("Done! Share this app link after you deploy (see README).")
else:
    st.info("Upload Excel files and click **Run Eligibility** in the sidebar.")
