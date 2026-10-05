"""
Capacity vs Order Distribution
Upload the Order Status Report (and last week's report to carry your routing) → download the report
"""

import base64
import datetime
import pathlib

import streamlit as st
from PIL import Image

from report_generator import generate_report

# ── Page config ────────────────────────────────────────────────────────────────
ASSETS_DIR = pathlib.Path(__file__).parent / "assets"
LOGO_PATH  = ASSETS_DIR / "jay_logo.jpg"
_logo_img  = Image.open(LOGO_PATH) if LOGO_PATH.exists() else "📦"

st.set_page_config(page_title="Capacity vs Order Distribution", page_icon=_logo_img, layout="wide")

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

    /* ── Metrics ──────────────────────────────────────────────────────── */
    [data-testid="stMetric"] {{
        background: {JAY_CHARCOAL};
        border: 1px solid #3a2f13;
        border-radius: 12px;
        padding: 0.7rem 0.9rem;
    }}
    [data-testid="stMetricValue"] {{
        color: {JAY_GOLD};
    }}

    hr {{ border-color: #3a2f13 !important; }}
</style>
""", unsafe_allow_html=True)

# ── Load reference workbook (bundled with the repo) ────────────────────────────
REF_PATH = pathlib.Path(__file__).parent / "reference_workbook.xlsx"


@st.cache_data(show_spinner=False)
def _load_ref():
    return REF_PATH.read_bytes()


ref_bytes = _load_ref()

# ── UI ─────────────────────────────────────────────────────────────────────────
_logo_b64 = base64.b64encode(LOGO_PATH.read_bytes()).decode() if LOGO_PATH.exists() else ""
st.markdown(f"""
<div class="jay-header">
    {f'<img src="data:image/jpeg;base64,{_logo_b64}" />' if _logo_b64 else ''}
    <div>
        <h1>Capacity vs Order Distribution</h1>
        <p>Upload the <b>Order Status Report</b> and click <b>Generate</b>. The report shows pending orders in
        containers by machine line, week by week for 12 weeks, against capacity (New capacity — Arul Sir), with
        <b>Excess Order</b> and <b>Short Order</b> under every TOTAL. Sheets: ORDER VS WEEK DISTRIBUTION (Zone →
        Region → Sub → Machine Line), AFRICA, AUSTRALIA &amp; EUROPE, USA (the masters and WORKING are hidden).</p>
    </div>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    st.markdown('<div class="jay-card"><h3>1 · Order Status Report</h3>', unsafe_allow_html=True)
    osr_file = st.file_uploader("Upload OSR (.xlsx)", type=["xlsx"], key="osr")
    st.markdown('</div>', unsafe_allow_html=True)
with col2:
    st.markdown('<div class="jay-card"><h3>2 · Last week\'s report</h3>', unsafe_allow_html=True)
    prev_file = st.file_uploader(
        "Optional — carries forward your MC ROUTING and Machines / Shifts", type=["xlsx"], key="prev")
    st.caption("Without it, the routing bundled with the app is used.")
    st.markdown('</div>', unsafe_allow_html=True)

with st.expander("⚙️ Options", expanded=False):
    as_of = st.date_input(
        "OSR date (As-Of)", value=datetime.date.today(),
        help="The first week column is the week that contains this date. Orders due earlier are counted in it.")

st.divider()
ready = osr_file is not None
btn = st.button("🚀 Generate Report", disabled=not ready, type="primary")
if not ready:
    st.info("Upload the Order Status Report to enable the Generate button.")

if btn and ready:
    with st.spinner("Building report…"):
        try:
            result_bytes, info = generate_report(
                osr_bytes=osr_file.read(),
                template_bytes=ref_bytes,
                as_of=as_of,
                prev_report_bytes=prev_file.read() if prev_file is not None else None,
            )
            filename = f"Capacity_vs_Order_Distribution_{as_of.strftime('%d%b%Y')}.xlsx"
            st.success("✅ Report generated successfully!")
            st.download_button(
                label="⬇️ Download Report",
                data=result_bytes,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

            m1, m2, m3 = st.columns(3)
            m1.metric("Weeks shown", f"W #{info['first_week']} – W #{info['last_week']}")
            m2.metric("Order rows / routed", f"{info['osr_rows']} / {info['routed_rows']}")
            m3.metric("Containers in report", f"{info['containers_in_report']:,.1f}")
            st.caption(
                f"Routing taken from the {info['routing_source']}. "
                f"A further {info['containers_after_last_week']:,.1f} containers are due after "
                f"W #{info['last_week']} and are not shown.")

            if info["inputs_changed"]:
                st.info(
                    "ℹ️ Machines / Shifts taken from your previous report (different from the app's defaults):\n\n"
                    + "\n".join(f"- {n}: {o[0]} machines × {o[1]} shifts → {w[0]} × {w[1]}"
                                for n, o, w in info["inputs_changed"]))
            if info["new_products"]:
                st.warning(
                    f"⚠️ **{len(info['new_products'])} product(s) with pending orders have no MC ROUTING and were not seen "
                    "before.** Unhide the WORKING sheet (right-click a tab → Unhide), type a route in column J, and they "
                    "flow into the report:\n\n"
                    + "\n".join(f"- {n}" for n in info["new_products"][:20])
                    + ("\n- …and more" if len(info["new_products"]) > 20 else ""))
            known_unrouted = [(n, c) for n, c in info["unrouted"] if n not in set(info["new_products"])]
            if known_unrouted:
                st.info(
                    f"ℹ️ {len(known_unrouted)} product(s) have pending orders but no route "
                    "(for example bulk lines) and are not counted:\n\n"
                    + "\n".join(f"- {n}: {c:,.0f} CFC" for n, c in known_unrouted[:10])
                    + ("\n- …and more" if len(known_unrouted) > 10 else ""))
            if info["unknown_routes"]:
                st.warning(
                    "⚠️ These routes are not in the routing table (hidden columns T:U of ORDER VS WEEK DISTRIBUTION), "
                    "so their orders are not counted:\n\n"
                    + "\n".join(f"- {r}: {c:,.0f} CFC" for r, c in info["unknown_routes"][:10]))
        except Exception as exc:
            st.error(f"❌ Error generating report:\n\n```\n{exc}\n```")
            raise

st.divider()
st.caption(
    "**Containers** = pending CFC ÷ CFC per container (item-specific where listed, else the machine line's). "
    "**Week** = the Monday of the order's requested date, or the first week if it is earlier. "
    "**Capacity** = unit capacity × Machines × Shifts (editable yellow cells in the report)."
)
st.markdown(
    f'<p style="text-align:center; color:{JAY_GOLD_DARK}; '
    f'font-size:0.78rem; letter-spacing:0.06em;">JAY · CAPACITY VS ORDER DISTRIBUTION</p>',
    unsafe_allow_html=True,
)
