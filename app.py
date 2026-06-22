"""
app.py
Streamlit UI for the UAT Excel Transformer.
"""

import re
from datetime import datetime

import streamlit as st

from transformer import transform

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="UAT Template Transformer",
    page_icon="📊",
    layout="centered",
)

# ── Styles ─────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .block-container { padding-top: 2rem; max-width: 780px; }
    .stAlert { border-radius: 8px; }
    div[data-testid="stFileUploader"] { border-radius: 8px; }
    .step-label {
        font-size: 0.78rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: #888;
        margin-bottom: 0.2rem;
    }
    .section-divider { margin: 1.5rem 0 1rem; border-top: 1px solid #eee; }
</style>
""", unsafe_allow_html=True)

# ── Header ─────────────────────────────────────────────────────────────────────
st.title("📊 UAT Template Transformer")
st.caption(
    "Upload your UAT tracking file and the strategy mapping template. "
    "The tool will add DT Location, Mapping Notes, and branch sheets automatically."
)

st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

# ── Step 1: Source systems ─────────────────────────────────────────────────────
st.markdown('<p class="step-label">Step 1 — Source Systems</p>', unsafe_allow_html=True)
st.markdown("Which source system(s) does this migration come **from**?")

col1, col2, col3 = st.columns(3)
with col1:
    use_andar    = st.checkbox("Andar",         value=True)
with col2:
    use_dtracker = st.checkbox("DTracker",      value=True)
with col3:
    use_re       = st.checkbox("Raiser's Edge", value=False)

source_systems = []
if use_andar:    source_systems.append("Andar")
if use_dtracker: source_systems.append("DTracker")
if use_re:       source_systems.append("Raiser's Edge")

if use_re:
    st.info(
        "ℹ️ Raiser's Edge support: the transformer will include RE fields "
        "if the mapping template contains a RE-labelled row. "
        "Branch sheets for RE will be added as **WRE1** and **WRE2**.",
        icon=None,
    )

st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

# ── Step 2: File uploads ───────────────────────────────────────────────────────
st.markdown('<p class="step-label">Step 2 — Upload Files</p>', unsafe_allow_html=True)

col_a, col_b = st.columns(2)

with col_a:
    st.markdown("**UAT Tracking File**")
    st.caption("e.g. `GAU_-_UWEM_Data_Migration_UAT_-_Round_1.xlsx`")
    uat_file = st.file_uploader(
        "UAT file", type=["xlsx"], label_visibility="collapsed", key="uat"
    )

with col_b:
    st.markdown("**Strategy Mapping Template**")
    st.caption("e.g. `ObjS01_05_wmarit_npsp__GeneralAccountingUnit__c_Template.xlsx`")
    template_file = st.file_uploader(
        "Template file", type=["xlsx"], label_visibility="collapsed", key="tmpl"
    )

st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

# ── Step 3: Output filename ────────────────────────────────────────────────────
st.markdown('<p class="step-label">Step 3 — Output File Name (optional)</p>', unsafe_allow_html=True)

# Auto-suggest a name based on UAT filename
suggested_name = ""
if uat_file:
    raw = uat_file.name.replace(".xlsx", "")
    # Replace UWEM → UWM and strip _temp suffixes
    suggested_name = re.sub(r"(?i)uwem", "UWM", raw)
    suggested_name = re.sub(r"(?i)_temp$", "", suggested_name)
    suggested_name += ".xlsx"

output_name = st.text_input(
    "Output filename",
    value=suggested_name,
    placeholder="UWM_GAU_-_Data_Migration_UAT_-_Round_1.xlsx",
    label_visibility="collapsed",
)
if output_name and not output_name.endswith(".xlsx"):
    output_name += ".xlsx"

st.markdown('<div class="section-divider"></div>', unsafe_allow_html=True)

# ── Step 4: Transform ──────────────────────────────────────────────────────────
st.markdown('<p class="step-label">Step 4 — Generate</p>', unsafe_allow_html=True)

ready = uat_file and template_file and len(source_systems) > 0

if not source_systems:
    st.warning("Please select at least one source system above.")

transform_btn = st.button(
    "⚙️  Generate Transformed Excel",
    disabled=not ready,
    use_container_width=True,
    type="primary",
)

if transform_btn and ready:
    progress_placeholder = st.empty()
    status_log = []

    def update_progress(msg):
        status_log.append(msg)
        progress_placeholder.markdown(
            "\n".join(f"- {m}" for m in status_log[-4:])
        )

    try:
        with st.spinner("Transforming…"):
            result_bytes = transform(
                uat_bytes=uat_file.read(),
                template_bytes=template_file.read(),
                source_systems=source_systems,
                progress_callback=update_progress,
            )

        progress_placeholder.empty()

        st.success("✅ Transformation complete! Click below to download.")

        final_name = output_name or f"UWM_UAT_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

        st.download_button(
            label=f"⬇️  Download  {final_name}",
            data=result_bytes,
            file_name=final_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        # Summary
        with st.expander("What was generated", expanded=True):
            st.markdown(f"""
- **Template sheet** renamed/updated with:
  - Andar Location column
  - DT Location column (populated from mapping template)
  - Mapping Notes column (with Andar Logic / DTracker Logic)
  - Calibri 12 font and thin borders throughout
- **Branch sheets created:**
  {"- WHALIF, WPEI (Andar branches)" if "Andar" in source_systems else ""}
  {"- WFREDE, WSTJOH (DTracker branches)" if "DTracker" in source_systems else ""}
  {"- WRE1, WRE2 (Raiser's Edge branches)" if "Raiser's Edge" in source_systems else ""}
- **Source systems included:** {", ".join(source_systems)}
""")

    except ValueError as e:
        progress_placeholder.empty()
        st.error(f"❌ {e}")
        st.markdown(
            "**Tip:** Make sure you've uploaded the correct files — "
            "the UAT tracking file first, and the strategy mapping template second."
        )
    except Exception as e:
        progress_placeholder.empty()
        st.error(f"❌ Unexpected error: {e}")
        st.markdown(
            "Please check both files are valid `.xlsx` files and try again. "
            "If the problem persists, contact your data team."
        )

# ── Footer ─────────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption(
    "UAT Template Transformer · United Way Centraide Canada · "
    "For issues contact your data migration team."
)
