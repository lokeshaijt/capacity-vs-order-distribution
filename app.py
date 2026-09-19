"""
MC Capacity vs Order — Week-Wise Report
Streamlit app: upload OSR + Production Register → download generated report
"""

import datetime
import io
import pathlib

import streamlit as st

from report_generator import generate_report

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="MC Capacity vs Order Report",
    page_icon="📦",
    layout="centered",
)

# ── Load reference workbook (bundled with the repo) ────────────────────────────
REF_PATH = pathlib.Path(__file__).parent / "reference_workbook.xlsx"

@st.cache_data(show_spinner=False)
def _load_ref():
    return REF_PATH.read_bytes()

ref_bytes = _load_ref()

# ── UI ─────────────────────────────────────────────────────────────────────────
st.title("📦 MC Capacity vs Order — Week-Wise Report")
st.markdown(
    "Upload the two source files and click **Generate Report** "
    "to download the updated workbook."
)

col1, col2 = st.columns(2)

with col1:
    st.subheader("1 · Order Status Report")
    osr_file = st.file_uploader(
        "Upload OrderStatusReport (.xlsx)",
        type=["xlsx"],
        key="osr",
    )

with col2:
    st.subheader("2 · Production Register")
    pr_file = st.file_uploader(
        "Upload Production Register (.xlsx)",
        type=["xlsx"],
        key="pr",
    )

st.divider()

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
