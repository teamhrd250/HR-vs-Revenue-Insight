# app.py
# HR vs Revenue Dashboard (Yearly 5Y template-ready)
# - Supports: yearly template HR_vs_Revenue_Template_Yearly_5Y.xlsx
# - Optional sheets: SALES_YR
# - Can also auto-detect monthly template if you later use it (REVENUE_MTH, HEADCOUNT_MTH, PAYROLL_MTH)

import os
from dataclasses import dataclass
from typing import Dict, List, Optional, Tuple

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st


# =========================
# UI / APP CONFIG
# =========================
st.set_page_config(
    page_title="HR vs Revenue Dashboard",
    page_icon="📊",
    layout="wide",
)

st.title("📊 HR vs Revenue Dashboard")
st.caption("Komparasi produktivitas manpower vs revenue (yearly / monthly auto-detect).")


# =========================
# CONSTANTS
# =========================
DEFAULT_TEMPLATE_PATH = os.path.join("data", "HR_vs_Revenue_Template_Yearly_5Y.xlsx")

YEARLY_SHEETS_REQUIRED = ["DIM_DEPARTMENT", "REVENUE_YR", "HEADCOUNT_YR", "PAYROLL_YR", "PARAMETERS"]
MONTHLY_SHEETS_REQUIRED = ["DIM_DEPARTMENT", "REVENUE_MTH", "HEADCOUNT_MTH", "PAYROLL_MTH", "PARAMETERS"]

# Yearly columns
REQ_COLS_YEARLY = {
    "DIM_DEPARTMENT": ["Dept_ID", "Dept_Name", "Function_Group", "Revenue_Driver_Flag"],
    "REVENUE_YR": ["Year", "Revenue_Recognized"],
    "HEADCOUNT_YR": ["Year", "Dept_ID", "Avg_Headcount"],
    "PAYROLL_YR": ["Year", "Dept_ID", "Payroll_Gross"],
    "PARAMETERS": ["Parameter", "Value"],
}

# Monthly columns
REQ_COLS_MONTHLY = {
    "DIM_DEPARTMENT": ["Dept_ID", "Dept_Name", "Function_Group", "Revenue_Driver_Flag"],
    "REVENUE_MTH": ["Month", "Revenue_Recognized"],
    "HEADCOUNT_MTH": ["Month", "Dept_ID", "Headcount_End"],
    "PAYROLL_MTH": ["Month", "Dept_ID", "Payroll_Gross"],
    "PARAMETERS": ["Parameter", "Value"],
}

# Optional sheets
OPTIONAL_SHEETS = ["SALES_YR", "SALES_MTH", "ATTRITION_EVENTS", "DERIVED_KPI_YR"]

VALID_FUNCTION_GROUPS = {"Sales", "Operations/Project", "Engineering", "Support", "Management"}
VALID_YN = {"Y", "N"}


# =========================
# HELPERS
# =========================
@dataclass
class LoadResult:
    mode: str  # "YEARLY" or "MONTHLY"
    sheets: Dict[str, pd.DataFrame]
    params: Dict[str, str]
    warnings: List[str]


def _safe_read_excel(xls: pd.ExcelFile, sheet: str) -> Optional[pd.DataFrame]:
    try:
        df = pd.read_excel(xls, sheet_name=sheet)
        # Normalize columns: strip spaces
        df.columns = [str(c).strip() for c in df.columns]
        return df
    except Exception:
        return None


def _as_str(v) -> str:
    if pd.isna(v):
        return ""
    return str(v).strip()


def _build_params(df_params: pd.DataFrame) -> Dict[str, str]:
    p = {}
    if df_params is None or df_params.empty:
        return p
    if not set(["Parameter", "Value"]).issubset(set(df_params.columns)):
        return p
    for _, row in df_params.iterrows():
        k = _as_str(row.get("Parameter"))
        v = row.get("Value")
        if k:
            p[k] = _as_str(v)
    return p


def _detect_mode(sheet_names: List[str]) -> str:
    s = set(sheet_names)
    if all(x in s for x in YEARLY_SHEETS_REQUIRED):
        return "YEARLY"
    if all(x in s for x in MONTHLY_SHEETS_REQUIRED):
        return "MONTHLY"
    # If mixed, prefer YEARLY if yearly core exists
    if {"DIM_DEPARTMENT", "PARAMETERS"}.issubset(s) and ("REVENUE_YR" in s or "HEADCOUNT_YR" in s or "PAYROLL_YR" in s):
        return "YEARLY"
    return "UNKNOWN"


def _validate_required_columns(df: pd.DataFrame, required_cols: List[str]) -> List[str]:
    missing = [c for c in required_cols if c not in df.columns]
    return missing


def _clean_dim_department(df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
    warnings = []
    d = df.copy()

    # Basic cleanup
    for c in ["Dept_ID", "Dept_Name", "Function_Group", "Revenue_Driver_Flag"]:
        if c in d.columns:
            d[c] = d[c].astype(str).str.strip()

    # Drop empty Dept_ID
    if "Dept_ID" in d.columns:
        before = len(d)
        d = d[d["Dept_ID"].astype(str).str.strip() != ""].copy()
        after = len(d)
        if after < before:
            warnings.append(f"DIM_DEPARTMENT: menghapus {before-after} baris Dept_ID kosong.")

    # Validate Function_Group
    if "Function_Group" in d.columns:
        bad_fg = sorted(set(d.loc[~d["Function_Group"].isin(VALID_FUNCTION_GROUPS), "Function_Group"]))
        if bad_fg:
            warnings.append(
                "DIM_DEPARTMENT: ada Function_Group di luar standar "
                f"{sorted(VALID_FUNCTION_GROUPS)} → ditemukan: {bad_fg}."
            )

    # Validate Revenue_Driver_Flag
    if "Revenue_Driver_Flag" in d.columns:
        bad_flag = sorted(set(d.loc[~d["Revenue_Driver_Flag"].isin(VALID_YN), "Revenue_Driver_Flag"]))
        if bad_flag:
            warnings.append("DIM_DEPARTMENT: Revenue_Driver_Flag harus Y/N → ditemukan: " + ", ".join(bad_flag))

    # Uniqueness Dept_ID
    if "Dept_ID" in d.columns:
        dup = d["Dept_ID"][d["Dept_ID"].duplicated()].unique().tolist()
        if dup:
            warnings.append(f"DIM_DEPARTMENT: Dept_ID duplikat ditemukan: {dup}. (Ini bikin join kacau.)")

    return d, warnings


def _coerce_year(df: pd.DataFrame, col: str) -> pd.DataFrame:
    out = df.copy()
    out[col] = pd.to_numeric(out[col], errors="coerce").astype("Int64")
    return out


def _coerce_month(df: pd.DataFrame, col: str) -> pd.DataFrame:
    out = df.copy()
    out[col] = pd.to_datetime(out[col], errors="coerce")
    return out


def _money(x) -> str:
    try:
        if x is None or (isinstance(x, float) and np.isnan(x)):
            return "-"
        x = float(x)
        return f"{x:,.0f}".replace(",", ".")
    except Exception:
        return "-"


def _pct(x) -> str:
    try:
        if x is None or (isinstance(x, float) and np.isnan(x)):
            return "-"
        x = float(x)
        return f"{x*100:.1f}%"
    except Exception:
        return "-"


@st.cache_data(show_spinner=False)
def load_workbook(file_bytes: Optional[bytes], default_path: str) -> LoadResult:
    warnings: List[str] = []

    if file_bytes is None:
        if not os.path.exists(default_path):
            return LoadResult("UNKNOWN", {}, {}, [f"File default tidak ditemukan: {default_path}"])
        xls = pd.ExcelFile(default_path)
    else:
        xls = pd.ExcelFile(file_bytes)

    sheet_names = xls.sheet_names
    mode = _detect_mode(sheet_names)
    if mode == "UNKNOWN":
        return LoadResult(
            "UNKNOWN",
            {},
            {},
            [f"Sheet tidak memenuhi template YEARLY atau MONTHLY. Sheet yang ada: {sheet_names}"],
        )

    # Read sheets
    sheets: Dict[str, pd.DataFrame] = {}
    # Required (based on mode)
    required = YEARLY_SHEETS_REQUIRED if mode == "YEARLY" else MONTHLY_SHEETS_REQUIRED
    required_cols = REQ_COLS_YEARLY if mode == "YEARLY" else REQ_COLS_MONTHLY

    for sh in required + OPTIONAL_SHEETS:
        df = _safe_read_excel(xls, sh)
        if df is not None:
            sheets[sh] = df

    # Validate required sheets exist
    for sh in required:
        if sh not in sheets:
            warnings.append(f"Sheet WAJIB tidak ditemukan: {sh}")

    if warnings:
        # still proceed as far as possible
        pass

    # Validate columns
    for sh in required:
        df = sheets.get(sh)
        if df is None:
            continue
        miss = _validate_required_columns(df, required_cols[sh])
        if miss:
            warnings.append(f"{sh}: kolom wajib hilang → {miss}")

    # Build params
    params = _build_params(sheets.get("PARAMETERS", pd.DataFrame()))

    # Clean DIM_DEPARTMENT
    if "DIM_DEPARTMENT" in sheets:
        dim, w = _clean_dim_department(sheets["DIM_DEPARTMENT"])
        sheets["DIM_DEPARTMENT"] = dim
        warnings.extend(w)

    return LoadResult(mode, sheets, params, warnings)


def kpi_cards(kpis: Dict[str, Tuple[str, str]], cols: int = 4):
    keys = list(kpis.keys())
    rows = (len(keys) + cols - 1) // cols
    i = 0
    for _ in range(rows):
        cc = st.columns(cols)
        for c in cc:
            if i >= len(keys):
                break
            k = keys[i]
            value, delta = kpis[k]
            c.metric(k, value, delta)
            i += 1


def calc_yearly(sheets: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    dim = sheets["DIM_DEPARTMENT"].copy()
    rev = sheets["REVENUE_YR"].copy()
    hc = sheets["HEADCOUNT_YR"].copy()
    pay = sheets["PAYROLL_YR"].copy()

    # Coerce Year
    rev = _coerce_year(rev, "Year")
    hc = _coerce_year(hc, "Year")
    pay = _coerce_year(pay, "Year")

    # Numeric coercions
    for c in ["Revenue_Recognized", "COGS_Direct"]:
        if c in rev.columns:
            rev[c] = pd.to_numeric(rev[c], errors="coerce").fillna(0.0)

    for c in ["Avg_Headcount", "Avg_FTE(optional)", "Avg_FTE", "New_Hires(optional)", "New_Hires", "Exits(optional)", "Exits"]:
        if c in hc.columns:
            hc[c] = pd.to_numeric(hc[c], errors="coerce")

    for c in ["Payroll_Gross", "Overtime(optional)", "Overtime", "Bonus(optional)", "Bonus", "Benefits(optional)", "Benefits", "Employer_Tax(optional)", "Employer_Tax", "Total_Manpower_Cost"]:
        if c in pay.columns:
            pay[c] = pd.to_numeric(pay[c], errors="coerce")

    # Build Total_Manpower_Cost if missing
    if "Total_Manpower_Cost" not in pay.columns:
        # sum available cost components
        cost_cols = [c for c in ["Payroll_Gross", "Overtime", "Bonus", "Benefits", "Employer_Tax"] if c in pay.columns]
        if cost_cols:
            pay["Total_Manpower_Cost"] = pay[cost_cols].sum(axis=1, skipna=True)
        else:
            pay["Total_Manpower_Cost"] = pay.get("Payroll_Gross", 0)

    # Join dim to headcount/payroll
    dim_small = dim[["Dept_ID", "Dept_Name", "Function_Group", "Revenue_Driver_Flag"]].copy()

    hc2 = hc.merge(dim_small, on="Dept_ID", how="left", suffixes=("", "_dim"))
    pay2 = pay.merge(dim_small, on="Dept_ID", how="left", suffixes=("", "_dim"))

    # Warnings for unmapped Dept_ID
    unmapped_hc = hc2[hc2["Function_Group"].isna()]["Dept_ID"].dropna().unique().tolist()
    unmapped_pay = pay2[pay2["Function_Group"].isna()]["Dept_ID"].dropna().unique().tolist()

    # Revenue totals by Year
    rev_tot = rev.groupby("Year", dropna=True, as_index=False).agg(
        Total_Revenue=("Revenue_Recognized", "sum"),
        Total_COGS=("COGS_Direct", "sum") if "COGS_Direct" in rev.columns else ("Revenue_Recognized", lambda x: 0.0),
    )
    rev_tot["Gross_Margin_Pct"] = np.where(
        rev_tot["Total_Revenue"] > 0,
        (rev_tot["Total_Revenue"] - rev_tot["Total_COGS"]) / rev_tot["Total_Revenue"],
        np.nan,
    )

    # Headcount totals by Year
    hc_tot = hc2.groupby("Year", dropna=True, as_index=False).agg(
        Total_Headcount=("Avg_Headcount", "sum"),
        RevenueDriver_Headcount=("Avg_Headcount", lambda x: np.nan),  # placeholder overwritten below
        Support_Headcount=("Avg_Headcount", lambda x: np.nan),
        Total_Hires=("New_Hires", "sum") if "New_Hires" in hc2.columns else ("Avg_Headcount", lambda x: np.nan),
        Total_Exits=("Exits", "sum") if "Exits" in hc2.columns else ("Avg_Headcount", lambda x: np.nan),
    )
    # fix driver/support breakdown
    if "Revenue_Driver_Flag" in hc2.columns:
        drv = hc2[hc2["Revenue_Driver_Flag"] == "Y"].groupby("Year", as_index=False).agg(
            RevenueDriver_Headcount=("Avg_Headcount", "sum")
        )
        sup = hc2[hc2["Revenue_Driver_Flag"] == "N"].groupby("Year", as_index=False).agg(
            Support_Headcount=("Avg_Headcount", "sum")
        )
        hc_tot = hc_tot.drop(columns=["RevenueDriver_Headcount", "Support_Headcount"]).merge(drv, on="Year", how="left").merge(
            sup, on="Year", how="left"
        )
    else:
        hc_tot["RevenueDriver_Headcount"] = np.nan
        hc_tot["Support_Headcount"] = np.nan

    # Payroll totals by Year
    pay_tot = pay2.groupby("Year", dropna=True, as_index=False).agg(
        Total_Manpower_Cost=("Total_Manpower_Cost", "sum"),
        Payroll_Gross=("Payroll_Gross", "sum") if "Payroll_Gross" in pay2.columns else ("Total_Manpower_Cost", "sum"),
    )

    # Master yearly KPI
    yr = rev_tot.merge(hc_tot, on="Year", how="outer").merge(pay_tot, on="Year", how="outer").sort_values("Year")
    yr["RPE"] = np.where(yr["Total_Headcount"] > 0, yr["Total_Revenue"] / yr["Total_Headcount"], np.nan)
    yr["MCR_Pct"] = np.where(yr["Total_Revenue"] > 0, yr["Total_Manpower_Cost"] / yr["Total_Revenue"], np.nan)

    # YoY growth
    yr["Revenue_YoY"] = yr["Total_Revenue"].pct_change()
    yr["Headcount_YoY"] = yr["Total_Headcount"].pct_change()
    yr["ManpowerCost_YoY"] = yr["Total_Manpower_Cost"].pct_change()
    yr["RPE_YoY"] = yr["RPE"].pct_change()
    yr["MCR_Delta"] = yr["MCR_Pct"].diff()

    # Breakdown by Function_Group (Headcount & Cost)
    hc_fg = hc2.groupby(["Year", "Function_Group"], dropna=False, as_index=False).agg(
        Headcount=("Avg_Headcount", "sum")
    )
    pay_fg = pay2.groupby(["Year", "Function_Group"], dropna=False, as_index=False).agg(
        Manpower_Cost=("Total_Manpower_Cost", "sum")
    )
    fg = hc_fg.merge(pay_fg, on=["Year", "Function_Group"], how="outer").sort_values(["Year", "Function_Group"])
    fg["Cost_per_HC"] = np.where(fg["Headcount"] > 0, fg["Manpower_Cost"] / fg["Headcount"], np.nan)

    # Business line breakdown (if exists)
    if "Business_Line" in rev.columns:
        rev_bl = rev.groupby(["Year", "Business_Line"], as_index=False).agg(Revenue=("Revenue_Recognized", "sum"))
    else:
        rev_bl = pd.DataFrame(columns=["Year", "Business_Line", "Revenue"])

    return {
        "yearly_kpi": yr,
        "fg_breakdown": fg,
        "rev_by_business_line": rev_bl,
        "unmapped_hc": pd.DataFrame({"Dept_ID": unmapped_hc}),
        "unmapped_pay": pd.DataFrame({"Dept_ID": unmapped_pay}),
    }


def calc_monthly(sheets: Dict[str, pd.DataFrame]) -> Dict[str, pd.DataFrame]:
    # This is a bonus: if you later use monthly template, app still runs.
    dim = sheets["DIM_DEPARTMENT"].copy()
    rev = sheets["REVENUE_MTH"].copy()
    hc = sheets["HEADCOUNT_MTH"].copy()
    pay = sheets["PAYROLL_MTH"].copy()

    rev = _coerce_month(rev, "Month")
    hc = _coerce_month(hc, "Month")
    pay = _coerce_month(pay, "Month")

    for c in ["Revenue_Recognized", "COGS_Direct"]:
        if c in rev.columns:
            rev[c] = pd.to_numeric(rev[c], errors="coerce").fillna(0.0)

    for c in ["Headcount_End", "FTE_End", "New_Hires", "Exits", "Avg_Headcount"]:
        if c in hc.columns:
            hc[c] = pd.to_numeric(hc[c], errors="coerce")

    for c in ["Payroll_Gross", "Overtime", "Bonus", "Benefits", "Employer_Tax", "Total_Manpower_Cost"]:
        if c in pay.columns:
            pay[c] = pd.to_numeric(pay[c], errors="coerce")

    if "Total_Manpower_Cost" not in pay.columns:
        cost_cols = [c for c in ["Payroll_Gross", "Overtime", "Bonus", "Benefits", "Employer_Tax"] if c in pay.columns]
        pay["Total_Manpower_Cost"] = pay[cost_cols].sum(axis=1, skipna=True) if cost_cols else pay.get("Payroll_Gross", 0)

    dim_small = dim[["Dept_ID", "Dept_Name", "Function_Group", "Revenue_Driver_Flag"]].copy()
    hc2 = hc.merge(dim_small, on="Dept_ID", how="left", suffixes=("", "_dim"))
    pay2 = pay.merge(dim_small, on="Dept_ID", how="left", suffixes=("", "_dim"))

    rev_tot = rev.groupby("Month", as_index=False).agg(
        Total_Revenue=("Revenue_Recognized", "sum"),
        Total_COGS=("COGS_Direct", "sum") if "COGS_Direct" in rev.columns else ("Revenue_Recognized", lambda x: 0.0),
    )
    rev_tot["Gross_Margin_Pct"] = np.where(
        rev_tot["Total_Revenue"] > 0,
        (rev_tot["Total_Revenue"] - rev_tot["Total_COGS"]) / rev_tot["Total_Revenue"],
        np.nan,
    )

    hc_tot = hc2.groupby("Month", as_index=False).agg(
        Total_Headcount=("Headcount_End", "sum"),
    )

    pay_tot = pay2.groupby("Month", as_index=False).agg(
        Total_Manpower_Cost=("Total_Manpower_Cost", "sum"),
    )

    m = rev_tot.merge(hc_tot, on="Month", how="outer").merge(pay_tot, on="Month", how="outer").sort_values("Month")
    m["RPE"] = np.where(m["Total_Headcount"] > 0, m["Total_Revenue"] / m["Total_Headcount"], np.nan)
    m["MCR_Pct"] = np.where(m["Total_Revenue"] > 0, m["Total_Manpower_Cost"] / m["Total_Revenue"], np.nan)

    # Monthly to yearly rollup (optional)
    m["Year"] = m["Month"].dt.year
    yr = m.groupby("Year", as_index=False).agg(
        Total_Revenue=("Total_Revenue", "sum"),
        Total_Headcount=("Total_Headcount", "mean"),  # approximate
        Total_Manpower_Cost=("Total_Manpower_Cost", "sum"),
        Gross_Margin_Pct=("Gross_Margin_Pct", "mean"),
    )
    yr["RPE"] = np.where(yr["Total_Headcount"] > 0, yr["Total_Revenue"] / yr["Total_Headcount"], np.nan)
    yr["MCR_Pct"] = np.where(yr["Total_Revenue"] > 0, yr["Total_Manpower_Cost"] / yr["Total_Revenue"], np.nan)

    return {"monthly_kpi": m, "yearly_kpi": yr}


def plot_line(df: pd.DataFrame, x: str, y: str, title: str):
    fig = px.line(df, x=x, y=y, markers=True)
    fig.update_layout(title=title, margin=dict(l=10, r=10, t=50, b=10))
    st.plotly_chart(fig, use_container_width=True)


def plot_bar(df: pd.DataFrame, x: str, y: str, color: Optional[str], title: str, barmode="group"):
    fig = px.bar(df, x=x, y=y, color=color, barmode=barmode)
    fig.update_layout(title=title, margin=dict(l=10, r=10, t=50, b=10))
    st.plotly_chart(fig, use_container_width=True)


def plot_area(df: pd.DataFrame, x: str, y: str, color: str, title: str):
    fig = px.area(df, x=x, y=y, color=color)
    fig.update_layout(title=title, margin=dict(l=10, r=10, t=50, b=10))
    st.plotly_chart(fig, use_container_width=True)


# =========================
# SIDEBAR: FILE INPUT
# =========================
with st.sidebar:
    st.header("⚙️ Input Data")
    st.write("Pakai **Upload Excel** atau default template dari repo.")

    uploaded = st.file_uploader("Upload file Excel (.xlsx)", type=["xlsx"])

    use_default = st.checkbox("Gunakan file default dari repo", value=(uploaded is None))
    if not use_default and uploaded is None:
        st.info("Upload Excel atau centang 'Gunakan file default'.")

    st.divider()
    st.header("🧭 Filter")
    st.write("Filter akan muncul setelah data berhasil dibaca.")


file_bytes = None
if uploaded is not None and not use_default:
    file_bytes = uploaded.getvalue()

load = load_workbook(file_bytes=file_bytes, default_path=DEFAULT_TEMPLATE_PATH)

if load.mode == "UNKNOWN":
    st.error("Tidak bisa membaca template. Cek sheet dan header kolom.")
    for w in load.warnings:
        st.warning(w)
    st.stop()

# Show warnings (non-blocking)
if load.warnings:
    with st.expander("⚠️ Data Warnings (klik untuk lihat)"):
        for w in load.warnings:
            st.warning(w)

params = load.params
currency = params.get("Currency_Code", "IDR")
company_name = params.get("Company_Name", "Perusahaan")
st.subheader(f"🏢 {company_name}  •  Mode: **{load.mode}**  •  Currency: **{currency}**")

# =========================
# COMPUTE KPIs
# =========================
if load.mode == "YEARLY":
    out = calc_yearly(load.sheets)
    yr = out["yearly_kpi"].copy()

    # Sidebar filter years
    with st.sidebar:
        years = yr["Year"].dropna().astype(int).unique().tolist()
        years = sorted(years)
        if years:
            yr_min, yr_max = min(years), max(years)
            year_range = st.slider("Rentang Tahun", min_value=yr_min, max_value=yr_max, value=(yr_min, yr_max))
        else:
            year_range = None

    if year_range is not None:
        yr = yr[(yr["Year"] >= year_range[0]) & (yr["Year"] <= year_range[1])].copy()

    if yr.empty:
        st.warning("Tidak ada data pada rentang tahun yang dipilih.")
        st.stop()

    # =========================
    # EXEC SUMMARY
    # =========================
    latest = yr.dropna(subset=["Year"]).sort_values("Year").iloc[-1]
    prev = yr.dropna(subset=["Year"]).sort_values("Year").iloc[-2] if len(yr) >= 2 else None

    def delta(a, b):
        if b is None or pd.isna(a) or pd.isna(b) or b == 0:
            return ""
        return f"{(a - b) / abs(b) * 100:.1f}%"

    kpis = {
        "Total Revenue (latest)": (_money(latest["Total_Revenue"]), delta(latest["Total_Revenue"], prev["Total_Revenue"] if prev is not None else None)),
        "Total Headcount (avg)": (f"{latest['Total_Headcount']:.1f}" if pd.notna(latest["Total_Headcount"]) else "-", delta(latest["Total_Headcount"], prev["Total_Headcount"] if prev is not None else None)),
        "RPE (Revenue/Employee)": (_money(latest["RPE"]), delta(latest["RPE"], prev["RPE"] if prev is not None else None)),
        "Manpower Cost Ratio (MCR)": (_pct(latest["MCR_Pct"]), f"{(latest['MCR_Delta']*100):+.1f} pp" if "MCR_Delta" in latest and pd.notna(latest["MCR_Delta"]) else ""),
        "Gross Margin": (_pct(latest["Gross_Margin_Pct"]), delta(latest["Gross_Margin_Pct"], prev["Gross_Margin_Pct"] if prev is not None else None)),
        "Total Manpower Cost": (_money(latest["Total_Manpower_Cost"]), delta(latest["Total_Manpower_Cost"], prev["Total_Manpower_Cost"] if prev is not None else None)),
    }

    st.markdown("### Executive Summary")
    kpi_cards(kpis, cols=3)

    # =========================
    # TRENDS
    # =========================
    c1, c2 = st.columns(2)
    with c1:
        plot_line(yr, "Year", "Total_Revenue", "Trend Revenue (Yearly)")
    with c2:
        plot_line(yr, "Year", "Total_Headcount", "Trend Headcount (Yearly)")

    c3, c4 = st.columns(2)
    with c3:
        plot_line(yr, "Year", "RPE", "Trend RPE (Revenue per Employee)")
    with c4:
        plot_line(yr, "Year", "MCR_Pct", "Trend MCR (Manpower Cost Ratio)")

    # Growth comparison
    st.markdown("### Growth: Revenue vs Headcount (YoY)")
    g = yr[["Year", "Revenue_YoY", "Headcount_YoY", "ManpowerCost_YoY"]].copy()
    g = g.melt(id_vars=["Year"], var_name="Metric", value_name="YoY")
    fig = px.line(g, x="Year", y="YoY", color="Metric", markers=True)
    fig.update_yaxes(tickformat=".0%")
    fig.update_layout(margin=dict(l=10, r=10, t=50, b=10))
    st.plotly_chart(fig, use_container_width=True)

    # =========================
    # STRUCTURE: FUNCTION GROUP
    # =========================
    st.markdown("### Struktur Manpower & Cost per Function Group")
    fg = out["fg_breakdown"].copy()

    # Filter to selected years
    if year_range is not None:
        fg = fg[(fg["Year"] >= year_range[0]) & (fg["Year"] <= year_range[1])].copy()

    # Handle missing Function_Group (unmapped dept)
    fg["Function_Group"] = fg["Function_Group"].fillna("UNMAPPED")

    c5, c6 = st.columns(2)
    with c5:
        plot_area(fg, "Year", "Headcount", "Function_Group", "Headcount by Function Group (Stacked)")
    with c6:
        plot_area(fg, "Year", "Manpower_Cost", "Function_Group", "Manpower Cost by Function Group (Stacked)")

    st.markdown("### Cost per Headcount per Function Group")
    plot_bar(fg, "Year", "Cost_per_HC", "Function_Group", "Cost per HC by Function Group", barmode="group")

    # =========================
    # REVENUE BY BUSINESS LINE
    # =========================
    rev_bl = out["rev_by_business_line"].copy()
    if not rev_bl.empty:
        st.markdown("### Revenue by Business Line")
        if year_range is not None:
            rev_bl = rev_bl[(rev_bl["Year"] >= year_range[0]) & (rev_bl["Year"] <= year_range[1])].copy()
        plot_area(rev_bl, "Year", "Revenue", "Business_Line", "Revenue Mix by Business Line (Stacked)")

    # =========================
    # DRIVER vs SUPPORT LENS
    # =========================
    st.markdown("### Revenue Driver vs Support (Lens)")
    lens = yr[["Year", "RevenueDriver_Headcount", "Support_Headcount", "Total_Headcount", "Total_Manpower_Cost", "Total_Revenue"]].copy()
    lens["Driver_Share_HC"] = np.where(lens["Total_Headcount"] > 0, lens["RevenueDriver_Headcount"] / lens["Total_Headcount"], np.nan)
    lens["Support_Share_HC"] = np.where(lens["Total_Headcount"] > 0, lens["Support_Headcount"] / lens["Total_Headcount"], np.nan)

    c7, c8 = st.columns(2)
    with c7:
        fig = px.bar(
            lens.melt(id_vars=["Year"], value_vars=["RevenueDriver_Headcount", "Support_Headcount"], var_name="Group", value_name="HC"),
            x="Year",
            y="HC",
            color="Group",
            barmode="stack",
            title="Headcount Split: Driver vs Support",
        )
        fig.update_layout(margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)

    with c8:
        fig = px.line(
            lens.melt(id_vars=["Year"], value_vars=["Driver_Share_HC", "Support_Share_HC"], var_name="Share", value_name="Pct"),
            x="Year",
            y="Pct",
            color="Share",
            markers=True,
            title="HC Share: Driver vs Support",
        )
        fig.update_yaxes(tickformat=".0%")
        fig.update_layout(margin=dict(l=10, r=10, t=50, b=10))
        st.plotly_chart(fig, use_container_width=True)

    # =========================
    # OPTIONAL: SALES PRODUCTIVITY
    # =========================
    if "SALES_YR" in load.sheets and not load.sheets["SALES_YR"].empty:
        st.markdown("### Sales Productivity (Optional)")
        sales = load.sheets["SALES_YR"].copy()
        sales.columns = [str(c).strip() for c in sales.columns]
        if "Year" in sales.columns:
            sales = _coerce_year(sales, "Year")
        for c in ["Deals_Closed", "Revenue_Booked", "Pipeline_Value(optional)", "Pipeline_Value", "Customer_Count(optional)", "Customer_Count"]:
            if c in sales.columns:
                sales[c] = pd.to_numeric(sales[c], errors="coerce")

        # normalize column names
        if "Pipeline_Value(optional)" in sales.columns and "Pipeline_Value" not in sales.columns:
            sales["Pipeline_Value"] = sales["Pipeline_Value(optional)"]
        if "Customer_Count(optional)" in sales.columns and "Customer_Count" not in sales.columns:
            sales["Customer_Count"] = sales["Customer_Count(optional)"]

        if year_range is not None:
            sales = sales[(sales["Year"] >= year_range[0]) & (sales["Year"] <= year_range[1])].copy()

        # KPI per year
        s_yr = sales.groupby("Year", as_index=False).agg(
            Deals=("Deals_Closed", "sum"),
            Revenue=("Revenue_Booked", "sum"),
            Salespeople=("Salesperson_ID", "nunique") if "Salesperson_ID" in sales.columns else ("Salesperson_Name", "nunique"),
        )
        s_yr["Revenue_per_Sales"] = np.where(s_yr["Salespeople"] > 0, s_yr["Revenue"] / s_yr["Salespeople"], np.nan)

        c9, c10 = st.columns(2)
        with c9:
            plot_line(s_yr, "Year", "Revenue_per_Sales", "Revenue per Salesperson (Yearly)")
        with c10:
            plot_line(s_yr, "Year", "Deals", "Deals Closed (Total)")

        st.markdown("#### Top/Bottom Sales (Latest Year)")
        latest_year = int(sales["Year"].dropna().max()) if sales["Year"].notna().any() else None
        if latest_year is not None:
            s_latest = sales[sales["Year"] == latest_year].copy()
            # ensure name
            if "Salesperson_Name" not in s_latest.columns and "Salesperson_ID" in s_latest.columns:
                s_latest["Salesperson_Name"] = s_latest["Salesperson_ID"]
            s_rank = s_latest.groupby("Salesperson_Name", as_index=False).agg(
                Revenue=("Revenue_Booked", "sum"),
                Deals=("Deals_Closed", "sum"),
            ).sort_values("Revenue", ascending=False)
            topn = st.slider("Top/Bottom N", 3, 20, 8, key="topn_sales")
            c11, c12 = st.columns(2)
            with c11:
                st.write("Top performers")
                st.dataframe(s_rank.head(topn), use_container_width=True, hide_index=True)
            with c12:
                st.write("Bottom performers")
                st.dataframe(s_rank.tail(topn).sort_values("Revenue", ascending=True), use_container_width=True, hide_index=True)

    # =========================
    # DATA QUALITY PANEL
    # =========================
    st.markdown("### Data Quality Checks")
    dq1, dq2 = st.columns(2)

    with dq1:
        st.write("Dept_ID unmapped (HEADCOUNT_YR)")
        um = out["unmapped_hc"]
        if um.empty:
            st.success("OK — semua Dept_ID di HEADCOUNT_YR ter-mapping.")
        else:
            st.warning("Ada Dept_ID di HEADCOUNT_YR yang tidak ada di DIM_DEPARTMENT.")
            st.dataframe(um.drop_duplicates(), use_container_width=True, hide_index=True)

    with dq2:
        st.write("Dept_ID unmapped (PAYROLL_YR)")
        um = out["unmapped_pay"]
        if um.empty:
            st.success("OK — semua Dept_ID di PAYROLL_YR ter-mapping.")
        else:
            st.warning("Ada Dept_ID di PAYROLL_YR yang tidak ada di DIM_DEPARTMENT.")
            st.dataframe(um.drop_duplicates(), use_container_width=True, hide_index=True)

    with st.expander("📄 Lihat tabel KPI tahunan (debug / export)"):
        show = yr.copy()
        # nicer formatting columns for display
        st.dataframe(show, use_container_width=True, hide_index=True)

else:
    # MONTHLY mode (if you later use monthly template)
    out = calc_monthly(load.sheets)
    m = out["monthly_kpi"].copy()
    yr = out["yearly_kpi"].copy()

    with st.sidebar:
        years = yr["Year"].dropna().astype(int).unique().tolist()
        years = sorted(years)
        year_range = st.slider("Rentang Tahun", min_value=min(years), max_value=max(years), value=(min(years), max(years))) if years else None

    if year_range is not None:
        m = m[(m["Month"].dt.year >= year_range[0]) & (m["Month"].dt.year <= year_range[1])].copy()
        yr = yr[(yr["Year"] >= year_range[0]) & (yr["Year"] <= year_range[1])].copy()

    st.markdown("### Executive Summary (Monthly dataset → rolled-up yearly view)")
    if yr.empty:
        st.warning("Tidak ada data.")
        st.stop()

    latest = yr.sort_values("Year").iloc[-1]
    prev = yr.sort_values("Year").iloc[-2] if len(yr) >= 2 else None

    def delta(a, b):
        if b is None or pd.isna(a) or pd.isna(b) or b == 0:
            return ""
        return f"{(a - b) / abs(b) * 100:.1f}%"

    kpis = {
        "Total Revenue (latest year)": (_money(latest["Total_Revenue"]), delta(latest["Total_Revenue"], prev["Total_Revenue"] if prev is not None else None)),
        "Avg Headcount (latest year)": (f"{latest['Total_Headcount']:.1f}" if pd.notna(latest["Total_Headcount"]) else "-", delta(latest["Total_Headcount"], prev["Total_Headcount"] if prev is not None else None)),
        "RPE (latest year)": (_money(latest["RPE"]), delta(latest["RPE"], prev["RPE"] if prev is not None else None)),
        "MCR (latest year)": (_pct(latest["MCR_Pct"]), ""),
    }
    kpi_cards(kpis, cols=4)

    c1, c2 = st.columns(2)
    with c1:
        plot_line(m, "Month", "Total_Revenue", "Monthly Revenue")
    with c2:
        plot_line(m, "Month", "Total_Headcount", "Monthly Headcount")

    c3, c4 = st.columns(2)
    with c3:
        plot_line(m, "Month", "RPE", "Monthly RPE")
    with c4:
        plot_line(m, "Month", "MCR_Pct", "Monthly MCR")

    st.markdown("### Yearly roll-up (from monthly)")
    c5, c6 = st.columns(2)
    with c5:
        plot_line(yr, "Year", "Total_Revenue", "Yearly Revenue (roll-up)")
    with c6:
        plot_line(yr, "Year", "RPE", "Yearly RPE (roll-up)")

    with st.expander("📄 Lihat tabel KPI monthly/yearly"):
        st.write("Monthly KPI")
        st.dataframe(m, use_container_width=True, hide_index=True)
        st.write("Yearly KPI (roll-up)")
        st.dataframe(yr, use_container_width=True, hide_index=True)


# =========================
# FOOTER NOTES
# =========================
st.divider()
st.caption(
    "Catatan: Untuk analisis 'Revenue per Function Group' yang benar-benar presisi, revenue idealnya punya dimensi org/dept "
    "(misal revenue booked per sales/dept, atau costed project team). Template ini fokus pada komparasi leverage manpower secara makro."
)
