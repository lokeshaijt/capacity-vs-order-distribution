"""
MC Capacity vs Order — Week-Wise Report
Streamlit app: upload OSR + Production Register → download generated report
"""

import base64
import datetime
import io
import pathlib

import streamlit as st
from PIL import Image

from report_generator import generate_report

# ── Page config ────────────────────────────────────────────────────────────────
ASSETS_DIR = pathlib.Path(__file__).parent / "assets"
LOGO_PATH  = ASSETS_DIR / "jay_logo.jpg"
_logo_img  = Image.open(LOGO_PATH) if LOGO_PATH.exists() else "📦"

st.set_page_config(
    page_title="MC Capacity vs Order Report",
    page_icon=_logo_img,
    layout="wide",
)

# ── JAY brand theme (black / gold, from the JAY logo) ───────────────────────────
JAY_GOLD        = "#F2B90C"
JAY_GOLD_LIGHT  = "#FFDE7A"
JAY_GOLD_DARK   = "#8A5A00"
JAY_BLACK       = "#0D0D0D"
JAY_CHARCOAL    = "#1A1A1A"
JAY_CREAM       = "#F5F1E6"

st.markdown(f"""
<style>
    .stApp {{
        background: radial-gradient(circle at 50% -20%, #262019 0%, {JAY_BLACK} 55%);
    }}
    .block-container {{
        max-width: 1100px;
        padding-top: 1.5rem;
    }}

    /* ── Brand header banner ─────────────────────────────────────────── */
    .jay-header {{
        display: flex;
        align-items: center;
        gap: 1.1rem;
        padding: 1.1rem 1.6rem;
        border-radius: 16px;
        margin-bottom: 1.6rem;
        background: linear-gradient(135deg, {JAY_CHARCOAL} 0%, #14110a 100%);
        border: 1px solid {JAY_GOLD_DARK};
        box-shadow: 0 0 24px rgba(242, 185, 12, 0.10);
    }}
    .jay-header img {{
        width: 56px; height: 56px; border-radius: 50%;
        border: 2px solid {JAY_GOLD};
    }}
    .jay-header h1 {{
        font-size: 1.5rem;
        margin: 0;
        background: linear-gradient(90deg, {JAY_GOLD_LIGHT}, {JAY_GOLD} 60%, {JAY_GOLD_DARK});
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 800;
    }}
    .jay-header p {{
        margin: 0.15rem 0 0 0;
        color: {JAY_CREAM};
        opacity: 0.75;
        font-size: 0.92rem;
    }}

    /* ── Section cards ───────────────────────────────────────────────── */
    .jay-card {{
        background: {JAY_CHARCOAL};
        border: 1px solid #3a2f13;
        border-radius: 14px;
        padding: 1.1rem 1.3rem 0.6rem 1.3rem;
        margin-bottom: 1.1rem;
    }}
    .jay-card h3 {{
        color: {JAY_GOLD};
        font-size: 1.02rem;
        margin-top: 0;
    }}

    /* ── Buttons ──────────────────────────────────────────────────────── */
    .stButton > button, .stDownloadButton > button {{
        background: linear-gradient(135deg, {JAY_GOLD_LIGHT}, {JAY_GOLD} 55%, {JAY_GOLD_DARK});
        color: #1a1200;
        font-weight: 700;
        border: none;
        border-radius: 10px;
        padding: 0.55rem 1.4rem;
        box-shadow: 0 2px 10px rgba(242, 185, 12, 0.25);
        transition: transform 0.05s ease-in-out, box-shadow 0.15s;
    }}
    .stButton > button:hover, .stDownloadButton > button:hover {{
        box-shadow: 0 4px 16px rgba(242, 185, 12, 0.45);
        transform: translateY(-1px);
        color: #1a1200;
    }}
    .stButton > button:disabled {{
        background: #3a3222;
        color: #8a8371;
        box-shadow: none;
    }}

    /* ── File uploader ───────────────────────────────────────────────── */
    [data-testid="stFileUploaderDropzone"] {{
        background: #14110a;
        border: 1.5px dashed {JAY_GOLD_DARK};
        border-radius: 12px;
    }}

    /* ── Expander (Options) ──────────────────────────────────────────── */
    [data-testid="stExpander"] {{
        background: {JAY_CHARCOAL};
        border: 1px solid #3a2f13;
        border-radius: 12px;
    }}

    hr {{ border-color: #3a2f13 !important; }}
</style>
""", unsafe_allow_html=True)

# ── UI ─────────────────────────────────────────────────────────────────────────
_logo_b64 = base64.b64encode(LOGO_PATH.read_bytes()).decode() if LOGO_PATH.exists() else ""
st.markdown(f"""
<div class="jay-header">
    {f'<img src="data:image/jpeg;base64,{_logo_b64}" />' if _logo_b64 else ''}
    <div>
        <h1>MC Capacity vs Order — Week-Wise Report</h1>
        <p>Upload the two source files and click <b>Generate Report</b> to download the updated workbook.</p>
    </div>
</div>
""", unsafe_allow_html=True)

# ── Load reference workbook (bundled with the repo) ────────────────────────────
REF_PATH = pathlib.Path(__file__).parent / "reference_workbook.xlsx"

@st.cache_data(show_spinner=False)
def _load_ref():
    return REF_PATH.read_bytes()

ref_bytes = _load_ref()

col1, col2 = st.columns(2)

with col1:
    st.markdown('<div class="jay-card"><h3>1 · Order Status Report</h3>', unsafe_allow_html=True)
    osr_file = st.file_uploader(
        "Upload OrderStatusReport (.xlsx)",
        type=["xlsx"],
        key="osr",
    )
    st.markdown('</div>', unsafe_allow_html=True)

with col2:
    st.markdown('<div class="jay-card"><h3>2 · Production Register</h3>', unsafe_allow_html=True)
    pr_file = st.file_uploader(
        "Upload Production Register (.xlsx)",
        type=["xlsx"],
        key="pr",
    )
    st.markdown('</div>', unsafe_allow_html=True)

# ── Options ────────────────────────────────────────────────────────────────────
with st.expander("⚙️ Options", expanded=False):
    c1, c2, c3 = st.columns(3)
    with c1:
        today_input = st.date_input(
            "As-Of Date (today)",
            value=datetime.date.today(),
        )
    with c2:
        n_forward = st.number_input(
            "Forward weeks to show",
            min_value=4, max_value=26, value=12, step=1,
        )
    with c3:
        n_achieved = st.number_input(
            "Past weeks for Achieved Capacity",
            min_value=2, max_value=8, value=4, step=1,
        )

# ── Generate ────────────────────────────────────────────────────────────────────
ready = osr_file is not None and pr_file is not None
generate_btn = st.button("🚀 Generate Report", disabled=not ready, type="primary")

if not ready:
    st.info("Please upload both files to enable the Generate button.")

if generate_btn and ready:
    with st.spinner("Building report — this takes a few seconds…"):
        try:
            result_bytes, skipped = generate_report(
                ref_bytes=ref_bytes,
                osr_bytes=osr_file.read(),
                pr_bytes=pr_file.read(),
                today=today_input,
                n_weeks_forward=int(n_forward),
                n_weeks_achieved=int(n_achieved),
            )

            filename = (
                f"MC_Capacity_vs_Order_"
                f"{today_input.strftime('%d-%b-%Y')}.xlsx"
            )

            st.success("✅ Report generated successfully!")

            st.download_button(
                label="⬇️ Download Report",
                data=result_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument"
                     ".spreadsheetml.sheet",
            )

            if skipped:
                skipped_names = sorted({s[2] for s in skipped
                                        if s[0] == "no_tbgs"})
                no_factors    = sorted({s[1] for s in skipped
                                        if s[0] == "no_factors"})
                if skipped_names:
                    st.warning(
                        f"⚠️ **{len(skipped_names)} product(s) in the Production Register "
                        f"could not be converted** (no TBGS/CTN conversion found). "
                        f"Their production is excluded from the Achieved Capacity sheet.\n\n"
                        + "\n".join(f"- {n}" for n in skipped_names[:20])
                        + ("\n- …and more" if len(skipped_names) > 20 else "")
                    )
                if no_factors:
                    st.warning(
                        f"⚠️ **{len(no_factors)} machine line(s) have no capacity factors** "
                        f"in MC MASTER — their rows are excluded.\n\n"
                        + "\n".join(f"- {n}" for n in no_factors[:10])
                    )

        except Exception as exc:
            st.error(f"❌ Error generating report:\n\n```\n{exc}\n```")
            raise

# ── Footer ─────────────────────────────────────────────────────────────────────
st.divider()
st.caption(
    "Reference data (MC MASTER, ITEM MASTER, TBGS PER CTN, routing layout) "
    "is bundled from the last saved workbook. "
    "To update master data, re-deploy with a new `reference_workbook.xlsx`."
)
st.markdown(
    f'<p style="text-align:center; color:{JAY_GOLD_DARK}; '
    f'font-size:0.78rem; letter-spacing:0.06em;">JAY · MC CAPACITY VS ORDER</p>',
    unsafe_allow_html=True,
)
