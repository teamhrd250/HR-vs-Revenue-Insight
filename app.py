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

# =========================
# HEADER (2 Logos + Title) — Board-grade layout
# - Logo perusahaan (utama) di kiri
# - Judul di tengah
# - Logo framework/solution di kanan (lebih kecil, secondary)
#
# Cara pakai:
# - Simpan logo perusahaan di: assets/logo_company.png  (atau .jpg/.jpeg)
# - Simpan logo framework/solution di: assets/logo_solution.png (atau .jpg/.jpeg)
# - Jika salah satu logo tidak ada, app tetap jalan (fallback otomatis).
# =========================
def _find_logo_paths() -> Dict[str, str]:
    """Cari logo di beberapa lokasi umum. Simpan file logo di repo agar ikut ter-deploy."""
    def pick(candidates):
        for p in candidates:
            if os.path.exists(p):
                return p
        return ""

    company = pick([
        os.path.join("assets", "logo_company.png"),
        os.path.join("assets", "logo_company.jpg"),
        os.path.join("assets", "logo_company.jpeg"),
        os.path.join("assets", "company_logo.png"),
        os.path.join("assets", "company_logo.jpg"),
        "logo_company.png",
        "company_logo.png",
    ])

    solution = pick([
        os.path.join("assets", "logo_solution.png"),
        os.path.join("assets", "logo_solution.jpg"),
        os.path.join("assets", "logo_solution.jpeg"),
        os.path.join("assets", "logo.png"),   # fallback: logo solusi lama
        os.path.join("assets", "logo.jpg"),
        os.path.join("assets", "logo.jpeg"),
        "Logo.png",
        "logo.png",
    ])

    return {"company": company, "solution": solution}


def _img_to_b64(logo_path: str) -> Tuple[str, str]:
    """Return (mime, base64) for an image file."""
    import base64, mimetypes
    mime, _ = mimetypes.guess_type(logo_path)
    mime = mime or "image/png"
    with open(logo_path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("utf-8")
    return mime, b64


def _render_header(logo_company: str, logo_solution: str) -> None:
    """Header premium: dua logo + hierarki judul rapih."""
    col_left, col_center, col_right = st.columns([1.6, 6.0, 1.6])

    # --- Logo perusahaan (utama)
    with col_left:
        if logo_company:
            mime, b64 = _img_to_b64(logo_company)
            st.markdown(
                f"""
                <div style="
                    background: rgba(255,255,255,0.06);
                    padding: 10px 12px;
                    border-radius: 14px;
                    display: inline-block;
                    box-shadow: 0 8px 22px rgba(0,0,0,0.25);
                ">
                    <img src="data:{mime};base64,{b64}" style="width: 155px; height: auto; display:block;" />
                </div>
                """,
                unsafe_allow_html=True,
            )
        else:
            # Fallback: jangan kosong total, tetap beri anchor yang rapi
            st.markdown(
                """
                <div style="
                    background: rgba(255,255,255,0.04);
                    padding: 12px 14px;
                    border-radius: 14px;
                    display: inline-block;
                ">
                    <div style="font-size:12px; color:#9aa0a6; line-height:1.2;">
                        Logo perusahaan<br/>belum di-set
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    # --- Judul & konteks
    with col_center:
        st.markdown(
            """
            <div style="padding-top:6px;">
                <div style="font-size:42px; font-weight:700; line-height:1.05; margin:0;">
                    HR vs Revenue Dashboard
                </div>
                <div style="color:#9aa0a6; font-size:14px; margin-top:8px;">
                    Board-Level Insight &amp; Risk Commentary
                </div>
                <div style="color:#7f8a96; font-size:13px; margin-top:8px;">
                    Komparasi produktivitas manpower vs revenue (yearly / monthly auto-detect).
                </div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    # --- Logo solution/framework (secondary)
    with col_right:
        if logo_solution:
            mime, b64 = _img_to_b64(logo_solution)
            st.markdown(
                f"""
                <div style="text-align:right; padding-top:14px; opacity:0.88;">
                    <img src="data:{mime};base64,{b64}" style="width: 105px; height:auto;" />
                    <div style="font-size:11px; color:#7f8a96; margin-top:6px;">
                        Analytics Framework
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        else:
            st.empty()

    st.markdown("---")


_logo_paths = _find_logo_paths()
_render_header(_logo_paths.get("company",""), _logo_paths.get("solution",""))


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



# =========================
# BOARD-LEVEL RISK COMMENTARY (Dynamic narrative)
# =========================
def _pp(x) -> str:
    """Format percentage point (pp). Input can be fraction (e.g., -0.019) or already pp."""
    try:
        if x is None or (isinstance(x, float) and np.isnan(x)):
            return "-"
        x = float(x)
        # If it's a fraction (abs <= 1), convert to pp
        if abs(x) <= 1.0:
            x = x * 100.0
        return f"{x:+.1f} pp"
    except Exception:
        return "-"


def board_commentary_yearly(
    kpi: Dict[str, float],
    dq: Optional[Dict] = None,
    smart_recs: Optional[List[Dict]] = None,
) -> Dict[str, List[str]]:
    """Generate board-level risk commentary (Bahasa Indonesia).

    Catatan desain:
    - Board commentary ini sengaja difokuskan pada "so what" (risiko, pertanyaan Board, arahan tindak lanjut).
    - Bagian "Smart Insights" di bawahnya fokus pada rekomendasi operasional/taktis.
    - Untuk menghindari pengulangan, fungsi ini bisa menerima smart_recs lalu melakukan de-dup sederhana.
    """
    dq = dq or {}
    score = dq.get("score")
    confidence = dq.get("confidence")

    rev = kpi.get("revenue_latest")
    rev_g = kpi.get("revenue_growth")
    hc = kpi.get("headcount_latest")
    hc_g = kpi.get("headcount_growth")
    rpe = kpi.get("rpe_latest")
    rpe_g = kpi.get("rpe_growth")
    mcr = kpi.get("mcr_latest")
    mcr_delta = kpi.get("mcr_delta")  # fraction diff (e.g., -0.019)
    gm = kpi.get("gross_margin_latest")
    mp_cost = kpi.get("manpower_cost_latest")
    mp_g = kpi.get("manpower_cost_growth")
    year = kpi.get("year_latest")

    # --- Data confidence disclaimer
    disclaimers: List[str] = []
    if isinstance(score, (int, float)) and score < 75:
        disclaimers.append(
            f"Kepercayaan insight: **{confidence}** (Data Quality Score {int(score)}/100). "
            "Beberapa metrik berpotensi bias jika payroll/headcount tidak lengkap per tahun/dept."
        )

    # --- Smart Insights alignment (anti-duplikasi)
    smart_recs = smart_recs or []

    def _norm_txt(x: str) -> str:
        return " ".join(str(x).lower().strip().split())

    smart_blob = " ".join(
        _norm_txt(" ".join([
            r.get("title", ""),
            r.get("why", ""),
            r.get("action", ""),
        ]))
        for r in smart_recs
    )

    smart_top = [r.get("title", "").strip() for r in smart_recs if str(r.get("title", "")).strip()][:3]

    alignment: List[str] = []
    if smart_top:
        alignment.append(
            "Smart Insights (taktis) di bawah ini sudah menyorot: "
            + ", ".join([f"**{t}**" for t in smart_top])
            + "."
        )
        alignment.append(
            "Board commentary ini fokus pada risiko strategis, pertanyaan governance, dan guardrail (bukan mengulang action item operasional)."
        )

    # --- Flags (rules)
    flags: List[str] = []
    if rev_g is not None and hc_g is not None and rpe_g is not None:
        if rev_g > 0.30 and hc_g < 0.05 and rpe_g > 0.30:
            flags.append("HIGH_LEVERAGE_GROWTH")
        if hc_g > rev_g + 0.05:
            flags.append("HC_OUTPACES_REVENUE")
        if rpe_g < -0.05:
            flags.append("RPE_DROP")

    if mcr is not None:
        if mcr < 0.06:
            flags.append("MCR_EXTREME_LOW")
        elif mcr < 0.10:
            flags.append("MCR_LOW")
        elif mcr > 0.30:
            flags.append("MCR_HIGH")

    if gm is not None and gm < 0.25:
        flags.append("MARGIN_PRESSURE")

    # --- Observations (facts)
    observations: List[str] = []
    y_lbl = str(int(year)) if year is not None and not pd.isna(year) else "latest"

    observations.append(
        f"Periode {y_lbl}: revenue **{_money(rev)}** ({_pct(rev_g)} YoY) dengan headcount rata-rata **{hc:.1f}** ({_pct(hc_g)} YoY)."
        if (rev is not None and hc is not None and pd.notna(hc))
        else "Periode terbaru menunjukkan perubahan signifikan pada metrik keuangan & manpower."
    )
    observations.append(f"RPE **{_money(rpe)}** ({_pct(rpe_g)} YoY) dan Gross Margin **{_pct(gm)}**.")
    observations.append(f"Biaya manpower **{_money(mp_cost)}** ({_pct(mp_g)} YoY) dengan MCR **{_pct(mcr)}** ({_pp(mcr_delta)} vs periode sebelumnya).")

    # --- Upside (konsekuensi baik) / Downside (risiko) / Guardrails
    upsides: List[str] = []
    downsides: List[str] = []
    guardrails: List[str] = []
    questions: List[str] = []
    actions: List[str] = []

    if "HIGH_LEVERAGE_GROWTH" in flags:
        upsides += [
            "Leverage organisasi membaik: pertumbuhan revenue didorong oleh peningkatan produktivitas, bukan ekspansi headcount besar.",
        ]
        guardrails += [
            "Pantau indikator kapasitas: overtime, backlog delivery, SLA breach, dan utilisasi per function (Sales/Ops/Engineering).",
            "Pantau indikator people risk: absenteeism, turnover (khusus role kunci), dan skor engagement (jika tersedia).",
        ]
        downsides += [
            "**Sustainability risk:** pertumbuhan revenue jauh melampaui pertumbuhan kapasitas; risiko over-reliance pada tenaga kerja eksisting.",
            "**Single-point-of-failure risk:** konsentrasi nilai per individu meningkat; churn pada role kunci dapat berdampak material pada delivery & revenue.",
        ]
        questions += [
            "Fungsi mana yang utilitasnya paling tinggi (Sales / Ops / Engineering) dan apa bottleneck utamanya?",
            "Leading indicator burnout apa yang dimonitor (overtime, absenteeism, engagement, turnover early warning)?",
        ]
        actions += [
            "Lakukan capacity vs demand review per function; identifikasi role bottleneck dan rencana penambahan kapasitas yang selektif.",
            "Perkuat succession coverage & knowledge transfer untuk posisi kunci (khususnya role yang memegang delivery/sales relationship).",
        ]

    if "MCR_EXTREME_LOW" in flags:
        upsides += [
            "Efisiensi biaya manpower sangat kuat (MCR sangat rendah), memberikan ruang investasi untuk pertumbuhan atau penguatan capability bila dikelola dengan guardrail.",
        ]
        guardrails += [
            "Tetapkan batas bawah (floor band) MCR yang selaras dengan target kualitas layanan; evaluasi jika MCR turun terlalu cepat tanpa penguatan sistem.",
            "Monitor leading indicator kualitas: rework, komplain pelanggan, cycle time delivery, dan incident rate.",
        ]
        downsides += [
            "**Cost structure risk:** MCR berada di level sangat rendah; bisa mencerminkan efisiensi kuat, namun juga potensi under-investment pada capability/coverage saat scaling.",
        ]
        questions += [
            "Penurunan MCR didorong oleh automation/proses (good) atau intensifikasi beban kerja manusia (risk)?",
            "Apakah perusahaan memiliki guardrail MCR (floor band) yang selaras dengan target kualitas layanan/SLA?",
        ]
        actions += [
            "Tetapkan guardrail MCR dan pantau bersama KPI delivery (SLA, customer satisfaction, rework) agar efisiensi tidak mengorbankan kualitas.",
            "Invest selektif pada tooling/automation dan capability building untuk menjaga leverage tanpa burnout.",
        ]

    if "HC_OUTPACES_REVENUE" in flags:
        downsides += [
            "**Scaling inefficiency risk:** headcount tumbuh lebih cepat dari revenue, berpotensi menekan RPE dan menaikkan fixed cost.",
        ]
        questions += [
            "Departemen mana yang mengalami ekspansi terbesar dan apakah ROI-nya terukur?",
        ]
        actions += [
            "Freeze hiring non-critical sementara; lakukan role rationalization dan realokasi kapasitas ke revenue driver.",
        ]

    if "RPE_DROP" in flags:
        downsides += [
            "**Productivity deterioration risk:** RPE turun; indikasi mismatch kapasitas vs demand atau penurunan efektivitas delivery/sales.",
        ]
        questions += [
            "Apakah penurunan RPE disebabkan oleh pricing/mix, idle capacity, atau execution issue?",
        ]
        actions += [
            "Lakukan productivity review per departemen (output vs cost) dan koreksi staffing model.",
        ]

    if "MCR_HIGH" in flags:
        downsides += [
            "**Margin compression risk:** MCR tinggi; struktur organisasi cenderung payroll-heavy sehingga margin rawan tertekan saat revenue melambat.",
        ]
        questions += [
            "Komponen cost apa yang menjadi driver (overtime/bonus/benefit/struktur gaji)?",
        ]
        actions += [
            "Audit komponen biaya manpower dan tetapkan approval hiring berbasis ROI (guardrail).",
        ]

    if "MARGIN_PRESSURE" in flags:
        downsides += [
            "**Margin quality risk:** gross margin rendah; potensi pricing pressure atau biaya delivery tidak terkendali.",
        ]
        questions += [
            "Project/customer mana yang memicu margin dilusi dan apa akar penyebabnya (scope creep, vendor, rework)?",
        ]
        actions += [
            "Perketat bid discipline dan lakukan post-mortem pada proyek low-margin.",
        ]

    if not downsides:
        downsides = ["Tidak ada red flag besar yang terdeteksi dari aturan dasar; tetap pantau konsistensi tren dan kelengkapan data."]
        questions = ["Apakah dataset payroll/headcount sudah lengkap dan konsisten ter-mapping Dept_ID untuk semua tahun yang dianalisis?"]
        actions = ["Lakukan review triwulan KPI (Revenue, RPE, MCR) + data validation checks untuk mencegah insight bias."]

    # --- Anti-duplikasi dengan Smart Insights (taktis)
    # Smart Insights di bawah sudah memberikan action item operasional.
    # Di sini kita pertahankan tindakan yang lebih bersifat governance/guardrail, dan buang yang terlalu mirip.
    if smart_blob and actions:
        dup_keywords = [
            "capacity", "productivity", "freeze", "audit", "bid", "post-mortem", "post mortem",
            "role rationalization", "role", "review per departemen", "bid discipline",
        ]

        def _looks_duplicate(a: str) -> bool:
            a_n = _norm_txt(a)
            return any((k in a_n) and (k in smart_blob) for k in dup_keywords)

        actions = [a for a in actions if not _looks_duplicate(a)] or actions

    return {
        "disclaimers": disclaimers,
        "alignment": alignment,
        "observations": observations,
        "risks": downsides,
        "upsides": upsides,
        "guardrails": guardrails,
        "questions": questions,
        "actions": actions,
    }





def board_commentary_paragraph_yearly(
    kpi: Dict[str, float],
    dq: Optional[Dict] = None,
    smart_recs: Optional[List[Dict]] = None,
) -> str:
    """Board-level narrative in 1 paragraph (Bahasa Indonesia), dinamis mengikuti KPI.

    Tujuan:
    - Menggabungkan fakta → konsekuensi (upside & downside) → mitigasi/guardrail dalam satu paragraf.
    - Tidak mengulang isi 'Smart Insights' (bagian itu taktis). Di sini fokus governance & risiko strategis.
    """
    dq = dq or {}
    smart_recs = smart_recs or []

    rev = kpi.get("revenue_latest")
    rev_g = kpi.get("revenue_growth")
    hc = kpi.get("headcount_latest")
    hc_g = kpi.get("headcount_growth")
    rpe = kpi.get("rpe_latest")
    rpe_g = kpi.get("rpe_growth")
    mcr = kpi.get("mcr_latest")
    mcr_delta = kpi.get("mcr_delta")
    gm = kpi.get("gross_margin_latest")
    mp_cost = kpi.get("manpower_cost_latest")
    mp_g = kpi.get("manpower_cost_growth")
    year = kpi.get("year_latest")

    y_lbl = str(int(year)) if year is not None and not pd.isna(year) else "terbaru"

    # Flags (rules sederhana)
    high_leverage = (rev_g is not None and hc_g is not None and rpe_g is not None and (rev_g > 0.30 and hc_g < 0.05 and rpe_g > 0.30))
    hc_outpaces_rev = (rev_g is not None and hc_g is not None and (hc_g > rev_g + 0.05))
    rpe_drop = (rpe_g is not None and (rpe_g < -0.05))
    mcr_extreme_low = (mcr is not None and (mcr < 0.06))
    mcr_high = (mcr is not None and (mcr > 0.30))
    margin_pressure = (gm is not None and (gm < 0.25))

    # Smart Insights titles (untuk sinkronisasi tanpa mengulang detail)
    smart_top = [str(r.get("title","")).strip() for r in smart_recs if str(r.get("title","")).strip()][:2]
    smart_hint = ""
    if smart_top:
        smart_hint = " Selaras dengan Smart Insights, area yang ikut disorot mencakup " + ", ".join([f"**{t}**" for t in smart_top]) + "."

    # Paragraph assembly (single markdown paragraph)
    para = (
        f"Pada periode **{y_lbl}**, revenue tercatat **{_money(rev)}** ({_pct(rev_g)} YoY) dengan headcount rata-rata **{hc:.1f}** ({_pct(hc_g)} YoY), "
        f"sehingga produktivitas (RPE) berada di **{_money(rpe)}** ({_pct(rpe_g)} YoY); di sisi biaya, manpower cost **{_money(mp_cost)}** ({_pct(mp_g)} YoY) "
        f"dengan MCR **{_pct(mcr)}** ({_pp(mcr_delta)} vs periode sebelumnya) dan gross margin **{_pct(gm)}**."
    )

    # Consequences (balanced)
    if high_leverage:
        para += (
            " Secara positif, pola ini menunjukkan **operational leverage** yang kuat: pertumbuhan revenue terutama datang dari peningkatan output per karyawan, bukan ekspansi headcount besar."
        )

    if hc_outpaces_rev:
        para += (
            " Namun, headcount yang tumbuh lebih cepat daripada revenue mengindikasikan **risiko scaling tidak efisien** (RPE rawan turun dan fixed cost meningkat)."
        )

    if rpe_drop:
        para += (
            " Penurunan RPE merupakan sinyal **penurunan produktivitas** (idle capacity, pricing/mix, atau isu eksekusi) yang perlu ditangani sebelum menjadi tren struktural."
        )

    if mcr_extreme_low:
        para += (
            " MCR yang sangat rendah memperlihatkan efisiensi biaya SDM yang kuat, tetapi juga menaikkan **risiko keberlanjutan kapasitas** bila efisiensi ini dicapai melalui intensifikasi beban kerja, "
            "ketergantungan pada individu kunci, dan potensi burnout/attrition (yang sering muncul dengan jeda waktu)."
        )
    elif mcr_high:
        para += (
            " MCR yang tinggi menandakan struktur organisasi cenderung payroll-heavy sehingga **margin lebih rentan tertekan** saat revenue melambat."
        )

    if margin_pressure:
        para += (
            " Gross margin yang rendah menambah **risiko kualitas margin** (pricing pressure atau biaya delivery tidak terkendali) sehingga kontrol scope, rework, dan disiplin komersial menjadi krusial."
        )

    # Mitigation / guardrails (governance-level, not tactical duplication)
    para += (
        " Untuk mitigasi, disarankan menetapkan **guardrail kapasitas & kualitas** (monitor utilisasi/overtime, absenteeism, turnover per role kritikal, backlog & SLA/incident), "
        "serta memperkuat **succession coverage dan knowledge transfer** pada posisi kunci agar pertumbuhan tetap scalable tanpa meningkatkan risiko operasional."
    )

    # Data confidence
    score = dq.get("score")
    conf = dq.get("confidence")
    if isinstance(score, (int, float)) and score < 75:
        para += f" Catatan: tingkat keyakinan insight saat ini **{conf}** (Data Quality Score {int(score)}/100), sehingga interpretasi perlu mempertimbangkan kelengkapan data payroll/headcount per tahun/dept."

    para += smart_hint
    return para

# =========================
# SMART INSIGHTS (Rule-based recommendations + diagnostics)
# =========================
def generate_recommendations_yearly(yr: pd.DataFrame, fg: pd.DataFrame) -> List[Dict]:
    """Rule-based recommendations for YEARLY mode."""
    recs: List[Dict] = []
    y = yr.dropna(subset=["Year"]).sort_values("Year").copy()
    if y.empty:
        return recs

    latest = y.iloc[-1]
    prev = y.iloc[-2] if len(y) >= 2 else None

    def add(sev: str, title: str, why: str, action: str, owner: str = "HR + Finance"):
        recs.append({"severity": sev, "title": title, "why": why, "action": action, "owner": owner})

    # Rule 1: Revenue vs Headcount growth mismatch
    if prev is not None and pd.notna(latest.get("Revenue_YoY")) and pd.notna(latest.get("Headcount_YoY")):
        rev_yoy = float(latest["Revenue_YoY"])
        hc_yoy = float(latest["Headcount_YoY"])
        if hc_yoy > rev_yoy + 0.05:
            add(
                "critical",
                "Headcount tumbuh lebih cepat daripada revenue",
                f"YoY Revenue: {rev_yoy:.1%} vs YoY Headcount: {hc_yoy:.1%}. Ini indikasi scaling tidak efisien (RPE berpotensi turun).",
                "Freeze hiring non-critical 1–2 kuartal, audit role & workload per fungsi, realokasi kapasitas dari support → revenue driver bila memungkinkan.",
                owner="HR + CEO/COO",
            )
        elif rev_yoy > hc_yoy + 0.10:
            add(
                "good",
                "Revenue tumbuh jauh lebih cepat dibanding headcount",
                f"YoY Revenue: {rev_yoy:.1%} vs YoY Headcount: {hc_yoy:.1%}. Leverage manpower membaik.",
                "Pertahankan disiplin headcount, invest selektif pada role bottleneck (sales/ops/engineering) + program retention untuk high performers.",
                owner="HR",
            )

    # Rule 2: RPE trend
    if prev is not None and pd.notna(latest.get("RPE")) and pd.notna(prev.get("RPE")):
        prev_rpe = float(prev["RPE"]) if float(prev["RPE"]) != 0 else float("nan")
        rpe_delta = (float(latest["RPE"]) - float(prev["RPE"])) / (abs(prev_rpe) + 1e-9)
        if rpe_delta < -0.05:
            add(
                "critical",
                "Revenue per Employee (RPE) turun signifikan",
                f"RPE turun {rpe_delta:.1%} dibanding tahun sebelumnya.",
                "Lakukan productivity review per departemen (output vs cost), hentikan rekrut role low-impact, dan perkuat capability (training terikat KPI).",
                owner="HR + Functional Heads",
            )
        elif rpe_delta > 0.05:
            add(
                "good",
                "RPE naik signifikan",
                f"RPE naik {rpe_delta:.1%}. Produktivitas manusia membaik.",
                "Scale yang sehat: jaga SOP, dokumentasi, dan pastikan kompensasi/insentif tetap kompetitif untuk mencegah attrition.",
                owner="HR",
            )

    # Rule 3: MCR threshold
    mcr = latest.get("MCR_Pct")
    if pd.notna(mcr):
        mcr = float(mcr)
        if mcr > 0.30:
            add(
                "critical",
                "Manpower Cost Ratio (MCR) terlalu tinggi (>30%)",
                f"MCR = {mcr:.1%}. Risiko margin tertekan; organisasi cenderung payroll-heavy.",
                "Audit komponen cost (overtime/bonus/benefit), optimasi staffing model per fungsi, dan terapkan guardrail hiring berbasis ROI.",
                owner="HR + Finance",
            )
        elif 0.20 < mcr <= 0.30:
            add(
                "warning",
                "MCR berada di zona waspada (20–30%)",
                f"MCR = {mcr:.1%}. Masih wajar, tapi perlu kontrol ketat saat revenue melambat.",
                "Monitoring overtime/bonus bulanan, approval hiring berlapis, dan evaluasi efektivitas insentif.",
                owner="HR + Finance",
            )
        elif mcr <= 0.20:
            add(
                "good",
                "MCR efisien (<20%)",
                f"MCR = {mcr:.1%}. Struktur biaya manpower relatif sehat.",
                "Jangan over-cut: fokus ke retention, capability building, dan automasi proses agar leverage tetap naik.",
                owner="HR",
            )

    # Rule 4: Driver vs Support share
    if pd.notna(latest.get("RevenueDriver_Headcount")) and pd.notna(latest.get("Support_Headcount")) and pd.notna(latest.get("Total_Headcount")):
        total = float(latest["Total_Headcount"]) if float(latest["Total_Headcount"]) != 0 else float("nan")
        if pd.notna(total):
            driver_share = float(latest["RevenueDriver_Headcount"]) / total
            support_share = float(latest["Support_Headcount"]) / total
            if support_share > 0.45:
                add(
                    "warning",
                    "Proporsi Support terlalu besar",
                    f"Support share = {support_share:.1%} (Driver share = {driver_share:.1%}). Jika revenue stagnan, ini sinyal organisasi 'gemuk'.",
                    "Tinjau proses support: automasi, shared service, simplifikasi workflow; pastikan pertumbuhan support selalu dikaitkan dengan growth revenue/driver.",
                    owner="COO + HR",
                )

    # Rule 5: Function group cost per HC anomaly (Support vs median)
    if fg is not None and not fg.empty and "Cost_per_HC" in fg.columns:
        fg_latest = fg[fg["Year"] == latest["Year"]].copy()
        if not fg_latest.empty:
            med = float(np.nanmedian(fg_latest["Cost_per_HC"].astype(float).values))
            sup = fg_latest[fg_latest["Function_Group"] == "Support"]
            if not sup.empty and pd.notna(med) and med > 0:
                sup_cph = sup["Cost_per_HC"].iloc[0]
                if pd.notna(sup_cph) and float(sup_cph) > 1.25 * med:
                    add(
                        "warning",
                        "Biaya per HC di Support lebih tinggi dari normal",
                        f"Support Cost/HC = {float(sup_cph):,.0f} vs median all groups = {med:,.0f}.",
                        "Review job leveling & grading, cek overlap role, dan tetapkan SLA/OKR support agar output terukur.",
                        owner="HR + Support Head",
                    )

    if not recs:
        add(
            "good",
            "Tidak ada red flag besar terdeteksi",
            "Aturan dasar tidak menemukan mismatch ekstrem pada revenue/headcount/cost.",
            "Lakukan review triwulan: RPE, MCR, dan driver-support ratio agar trend negatif tertangkap lebih awal.",
            owner="HR",
        )

    priority = {"critical": 0, "warning": 1, "good": 2}
    return sorted(recs, key=lambda r: priority.get(r.get("severity", "warning"), 9))


def data_quality_score_yearly(
    yr: pd.DataFrame,
    dim: pd.DataFrame,
    hc: pd.DataFrame,
    pay: pd.DataFrame,
    rev: pd.DataFrame,
) -> Dict:
    """Score 0-100 + confidence + issues."""
    issues: List[str] = []
    score = 100

    years = sorted(yr["Year"].dropna().astype(int).unique().tolist()) if "Year" in yr.columns else []
    if len(years) < 3:
        score -= 20
        issues.append("Data tahun < 3 → tren YoY kurang reliabel.")
    elif len(years) < 5:
        score -= 8
        issues.append("Data < 5 tahun → komparasi terbatas.")

    for col, penalty in [("Total_Revenue", 15), ("Total_Headcount", 15), ("Total_Manpower_Cost", 15)]:
        if col not in yr.columns or yr[col].isna().mean() > 0.2:
            score -= penalty
            issues.append(f"Banyak nilai kosong pada {col} → KPI bisa bias.")

    if "Dept_ID" in hc.columns and "Dept_ID" in dim.columns:
        unmapped_hc = set(hc["Dept_ID"].astype(str)) - set(dim["Dept_ID"].astype(str))
        unmapped_hc = {x for x in unmapped_hc if x.strip() != ""}
        if len(unmapped_hc) > 0:
            score -= min(20, 3 * len(unmapped_hc))
            issues.append(f"Dept_ID HEADCOUNT tidak ter-mapping di DIM (contoh): {sorted(list(unmapped_hc))[:8]}")

    if "Dept_ID" in pay.columns and "Dept_ID" in dim.columns:
        unmapped_pay = set(pay["Dept_ID"].astype(str)) - set(dim["Dept_ID"].astype(str))
        unmapped_pay = {x for x in unmapped_pay if x.strip() != ""}
        if len(unmapped_pay) > 0:
            score -= min(20, 3 * len(unmapped_pay))
            issues.append(f"Dept_ID PAYROLL tidak ter-mapping di DIM (contoh): {sorted(list(unmapped_pay))[:8]}")

    if "Revenue_Recognized" in rev.columns:
        neg = (pd.to_numeric(rev["Revenue_Recognized"], errors="coerce") < 0).sum()
        if neg > 0:
            score -= min(15, 5 * int(neg))
            issues.append(f"Ada revenue negatif ({int(neg)} baris). Pastikan refund/credit note valid.")

    score = max(0, min(100, score))
    confidence = "High" if score >= 85 else ("Medium" if score >= 65 else "Low")
    return {"score": score, "confidence": confidence, "issues": issues, "years": years}


def dept_root_cause_yearly(
    year: int,
    hc_raw: pd.DataFrame,
    pay_raw: pd.DataFrame,
    dim: pd.DataFrame,
) -> pd.DataFrame:
    """Dept-level root cause table for a given year."""
    h = hc_raw.copy()
    p = pay_raw.copy()
    d = dim[["Dept_ID", "Dept_Name", "Function_Group", "Revenue_Driver_Flag"]].copy()

    h["Year"] = pd.to_numeric(h.get("Year"), errors="coerce").astype("Int64")
    p["Year"] = pd.to_numeric(p.get("Year"), errors="coerce").astype("Int64")

    h["Avg_Headcount"] = pd.to_numeric(h.get("Avg_Headcount"), errors="coerce")
    if "Total_Manpower_Cost" in p.columns:
        p["Total_Manpower_Cost"] = pd.to_numeric(p.get("Total_Manpower_Cost"), errors="coerce")
    else:
        p["Total_Manpower_Cost"] = pd.to_numeric(p.get("Payroll_Gross"), errors="coerce")

    h = h[h["Year"] == year].copy()
    p = p[p["Year"] == year].copy()

    h_agg = h.groupby("Dept_ID", as_index=False).agg(Headcount=("Avg_Headcount", "sum"))
    p_agg = p.groupby("Dept_ID", as_index=False).agg(Manpower_Cost=("Total_Manpower_Cost", "sum"))

    out = h_agg.merge(p_agg, on="Dept_ID", how="outer").merge(d, on="Dept_ID", how="left")
    out["Cost_per_HC"] = np.where(out["Headcount"] > 0, out["Manpower_Cost"] / out["Headcount"], np.nan)
    return out.sort_values("Manpower_Cost", ascending=False)


def anomaly_flags_yearly(yr: pd.DataFrame) -> List[Dict]:
    """Detect unusual YoY spikes and contradictions."""
    flags: List[Dict] = []
    y = yr.dropna(subset=["Year"]).sort_values("Year").copy()
    if len(y) < 3:
        return flags

    latest = y.iloc[-1]

    if pd.notna(latest.get("Headcount_YoY")) and float(latest["Headcount_YoY"]) > 0.25:
        flags.append({
            "severity": "warning",
            "title": "Lonjakan headcount tidak normal",
            "why": f"Headcount YoY = {float(latest['Headcount_YoY']):.1%} (>25%).",
            "action": "Audit hiring: role yang ditambah, justification ROI, dan dampak ke RPE.",
            "owner": "HR + COO"
        })

    if pd.notna(latest.get("ManpowerCost_YoY")) and float(latest["ManpowerCost_YoY"]) > 0.30:
        flags.append({
            "severity": "warning",
            "title": "Lonjakan biaya manpower tidak normal",
            "why": f"Manpower Cost YoY = {float(latest['ManpowerCost_YoY']):.1%} (>30%).",
            "action": "Breakdown overtime/bonus/benefit; cek perubahan struktur gaji atau reclass cost.",
            "owner": "HR + Finance"
        })

    if pd.notna(latest.get("Revenue_YoY")) and pd.notna(latest.get("Headcount_YoY")):
        if float(latest["Revenue_YoY"]) < 0 and float(latest["Headcount_YoY"]) > 0.10:
            flags.append({
                "severity": "critical",
                "title": "Revenue turun tapi headcount tetap naik",
                "why": f"Revenue YoY = {float(latest['Revenue_YoY']):.1%}, Headcount YoY = {float(latest['Headcount_YoY']):.1%}.",
                "action": "Immediate hiring freeze + reallocation kapasitas + evaluasi produktivitas per function.",
                "owner": "CEO + HR + Finance"
            })

    priority = {"critical": 0, "warning": 1, "good": 2}
    return sorted(flags, key=lambda r: priority.get(r.get("severity", "warning"), 9))


def smart_recommendations_yearly(
    yr: pd.DataFrame,
    fg: pd.DataFrame,
    dim: pd.DataFrame,
    hc_raw: pd.DataFrame,
    pay_raw: pd.DataFrame,
    rev_raw: pd.DataFrame,
) -> Dict:
    """Combine recommendations + anomalies + data quality + dept root-cause."""
    base = generate_recommendations_yearly(yr, fg)
    anom = anomaly_flags_yearly(yr)

    dq = data_quality_score_yearly(
        yr=yr,
        dim=dim,
        hc=hc_raw,
        pay=pay_raw,
        rev=rev_raw,
    )

    latest_year = int(yr.dropna(subset=["Year"]).sort_values("Year").iloc[-1]["Year"])
    dept_tbl = dept_root_cause_yearly(latest_year, hc_raw, pay_raw, dim)

    all_recs = base + anom
    priority = {"critical": 0, "warning": 1, "good": 2}
    all_recs = sorted(all_recs, key=lambda r: priority.get(r.get("severity", "warning"), 9))

    return {"dq": dq, "recs": all_recs, "dept_tbl": dept_tbl, "latest_year": latest_year}


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

# Ringkasan singkat (auto) — supaya pembaca langsung dapat konteks sebelum drill-down
st.caption(
    "Ringkasan: dashboard ini membandingkan leverage manpower terhadap pertumbuhan revenue. "
    "Untuk pembacaan yang aman, interpretasikan MCR selalu bersama Revenue Growth, RPE, dan Headcount."
)

# Board-level risk commentary (full) — ditampilkan sebagai expander (detail)
dengan_smart = smart_recommendations_yearly(
    yr=yr,
    fg=out["fg_breakdown"],
    dim=load.sheets["DIM_DEPARTMENT"],
    hc_raw=load.sheets["HEADCOUNT_YR"],
    pay_raw=load.sheets["PAYROLL_YR"],
    rev_raw=load.sheets["REVENUE_YR"],
)

dq_now = data_quality_score_yearly(
    yr=yr,
    dim=load.sheets["DIM_DEPARTMENT"],
    hc=load.sheets["HEADCOUNT_YR"],
    pay=load.sheets["PAYROLL_YR"],
    rev=load.sheets["REVENUE_YR"],
)

kpi_num = {
    "year_latest": float(latest["Year"]) if pd.notna(latest.get("Year")) else np.nan,
    "revenue_latest": float(latest["Total_Revenue"]) if pd.notna(latest.get("Total_Revenue")) else np.nan,
    "revenue_growth": float(latest["Revenue_YoY"]) if pd.notna(latest.get("Revenue_YoY")) else np.nan,
    "headcount_latest": float(latest["Total_Headcount"]) if pd.notna(latest.get("Total_Headcount")) else np.nan,
    "headcount_growth": float(latest["Headcount_YoY"]) if pd.notna(latest.get("Headcount_YoY")) else np.nan,
    "rpe_latest": float(latest["RPE"]) if pd.notna(latest.get("RPE")) else np.nan,
    "rpe_growth": float(latest["RPE_YoY"]) if pd.notna(latest.get("RPE_YoY")) else np.nan,
    "mcr_latest": float(latest["MCR_Pct"]) if pd.notna(latest.get("MCR_Pct")) else np.nan,
    "mcr_delta": float(latest["MCR_Delta"]) if pd.notna(latest.get("MCR_Delta")) else np.nan,
    "gross_margin_latest": float(latest["Gross_Margin_Pct"]) if pd.notna(latest.get("Gross_Margin_Pct")) else np.nan,
    "manpower_cost_latest": float(latest["Total_Manpower_Cost"]) if pd.notna(latest.get("Total_Manpower_Cost")) else np.nan,
    "manpower_cost_growth": float(latest["ManpowerCost_YoY"]) if pd.notna(latest.get("ManpowerCost_YoY")) else np.nan,
}

board = board_commentary_yearly(kpi=kpi_num, dq=dq_now, smart_recs=dengan_smart.get("recs"))

with st.expander("📌 Board-Level Risk Commentary (Lengkap)", expanded=False):
    st.markdown(board_commentary_paragraph_yearly(kpi=kpi_num, dq=dq_now, smart_recs=dengan_smart.get("recs")))

# =========================
# SMART INSIGHTS & REKOMENDASI OTOMATIS
# =========================
st.markdown("### 🧠 Smart Insights & Rekomendasi Otomatis")

# Re-use hasil diagnostik yang sama agar narasi Board-level & Smart Insights sinkron
smart = dengan_smart

dq = smart["dq"]
recs = smart["recs"]
dept_tbl = smart["dept_tbl"]
latest_year_smart = smart["latest_year"]

cqa1, cqa2 = st.columns([2, 5])
with cqa1:
    st.metric("Data Quality Score", f"{dq['score']}/100", dq["confidence"])
with cqa2:
    if dq["issues"]:
        st.warning("Beberapa isu kualitas data terdeteksi. Insight tetap jalan, tapi confidence bisa turun.")
        with st.expander("Lihat isu kualitas data"):
            for it in dq["issues"]:
                st.write(f"- {it}")
    else:
        st.success("Kualitas data bagus. Insight lebih bisa dipercaya.")

for r in recs:
    sev = r.get("severity", "warning")
    msg = f"**{r.get('title','')}** — {r.get('why','')}"
    if sev == "critical":
        st.error(msg)
    elif sev == "warning":
        st.warning(msg)
    else:
        st.success(msg)

with st.expander("📌 Action Plan (detail)"):
    df_recs = pd.DataFrame(recs)[["severity", "title", "why", "action", "owner"]]
    st.dataframe(df_recs, use_container_width=True, hide_index=True)

st.markdown(f"### 🔎 Root Cause (Dept-level) — {latest_year_smart}")

focus_cols = ["Dept_ID", "Dept_Name", "Function_Group", "Revenue_Driver_Flag", "Headcount", "Manpower_Cost", "Cost_per_HC"]
dept_view = dept_tbl.copy()
dept_view = dept_view[[c for c in focus_cols if c in dept_view.columns]].copy()

top_n = st.slider("Top N Dept untuk analisa", 5, 30, 10, key="dept_topn")
c1r, c2r, c3r = st.columns(3)

with c1r:
    st.write("Top Manpower Cost")
    st.dataframe(dept_view.sort_values("Manpower_Cost", ascending=False).head(top_n), use_container_width=True, hide_index=True)

with c2r:
    st.write("Top Headcount")
    st.dataframe(dept_view.sort_values("Headcount", ascending=False).head(top_n), use_container_width=True, hide_index=True)

with c3r:
    st.write("Highest Cost per HC")
    st.dataframe(dept_view.sort_values("Cost_per_HC", ascending=False).head(top_n), use_container_width=True, hide_index=True)

st.markdown("#### Scatter: Headcount vs Manpower Cost (Dept)")
if "Headcount" in dept_tbl.columns and "Manpower_Cost" in dept_tbl.columns:
    fig = px.scatter(
        dept_tbl,
        x="Headcount",
        y="Manpower_Cost",
        color="Function_Group",
        hover_data=["Dept_ID", "Dept_Name", "Revenue_Driver_Flag"],
        title="Dept Cost Structure: mana yang besar & mahal",
    )
    fig.update_layout(margin=dict(l=10, r=10, t=50, b=10))
    st.plotly_chart(fig, use_container_width=True)


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