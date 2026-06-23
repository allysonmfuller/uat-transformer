"""
app.py
Streamlit UI for the UAT Template Generator.
User uploads ONE file (the strategy mapping template).
The blank UAT template is hardcoded inside this app.
"""

import base64
import re
from datetime import datetime
from pathlib import Path

import streamlit as st

from mapping_parser import parse_mapping_template
from template_builder import build_output

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="UAT Template Generator",
    page_icon="📊",
    layout="centered",
)

st.markdown("""
<style>
    .block-container { padding-top: 2rem; max-width: 780px; }
    .step-label {
        font-size: 0.78rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: #888;
        margin-bottom: 0.25rem;
    }
    .divider { margin: 1.5rem 0 1rem; border-top: 1px solid #eee; }
</style>
""", unsafe_allow_html=True)

# ── Load blank template from file ─────────────────────────────────────────────
# The blank template lives in the repo alongside app.py
BLANK_TEMPLATE_PATH = Path(__file__).parent / "UWC_NPSP_User_Acceptance_Testing_Template.xlsx"

@st.cache_data
def load_blank_template() -> bytes:
    with open(BLANK_TEMPLATE_PATH, "rb") as f:
        return f.read()

try:
    blank_bytes = load_blank_template()
except FileNotFoundError:
    st.error(
        "❌ Blank template file not found. "
        "Make sure `UWC_NPSP_User_Acceptance_Testing_Template.xlsx` "
        "is in the same folder as `app.py`."
    )
    st.stop()

# ── Header ─────────────────────────────────────────────────────────────────────
st.title("📊 UAT Template Generator")
st.caption(
    "Upload the strategy mapping template for any object. "
    "The app generates a fully formatted UAT tracking file — "
    "no UAT file needed."
)

st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

# ── Step 1: Source systems ─────────────────────────────────────────────────────
st.markdown('<p class="step-label">Step 1 — Source Systems</p>', unsafe_allow_html=True)
st.markdown("Which source system(s) is this migration coming **from**?")

col1, col2, col3 = st.columns(3)
with col1:
    use_andar = st.checkbox("Andar",         value=True)
with col2:
    use_dt    = st.checkbox("DTracker",      value=True)
with col3:
    use_re    = st.checkbox("Raiser's Edge", value=False)

source_systems = []
if use_andar: source_systems.append("Andar")
if use_dt:    source_systems.append("DTracker")
if use_re:    source_systems.append("Raiser's Edge")

if not source_systems:
    st.warning("Please select at least one source system.")

st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

# ── Step 2: Upload mapping template ───────────────────────────────────────────
st.markdown('<p class="step-label">Step 2 — Upload Strategy Mapping Template</p>',
            unsafe_allow_html=True)
st.caption(
    "e.g. `ObjS02_02_wstrat_npe01__OppPayment__c_Template.xlsx` — "
    "works for any object type."
)

mapping_file = st.file_uploader(
    "Mapping template", type=["xlsx"], label_visibility="collapsed"
)

st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

# ── Step 3: Output filename ────────────────────────────────────────────────────
st.markdown('<p class="step-label">Step 3 — Output File Name (optional)</p>',
            unsafe_allow_html=True)

suggested_name = ""
if mapping_file:
    raw = mapping_file.name.replace(".xlsx", "")
    # Strip strategy author prefix (wstrat, wmarit, etc.) and Template suffix
    raw = re.sub(r"_w[a-z]+_", "_", raw)
    raw = re.sub(r"_Template$", "", raw, flags=re.IGNORECASE)
    raw = re.sub(r"_temp$",     "", raw, flags=re.IGNORECASE)
    suggested_name = f"UWM_{raw}_UAT_Round_1.xlsx"

output_name = st.text_input(
    "Output filename",
    value=suggested_name,
    placeholder="UWM_Payment_UAT_Round_1.xlsx",
    label_visibility="collapsed",
)
if output_name and not output_name.endswith(".xlsx"):
    output_name += ".xlsx"

st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

# ── Step 4: Generate ──────────────────────────────────────────────────────────
st.markdown('<p class="step-label">Step 4 — Generate</p>', unsafe_allow_html=True)

ready = mapping_file is not None and len(source_systems) > 0

generate_btn = st.button(
    "⚙️  Generate UAT Template",
    disabled=not ready,
    use_container_width=True,
    type="primary",
)

if generate_btn and ready:
    progress_placeholder = st.empty()
    log_lines = []

    def update_progress(msg):
        log_lines.append(msg)
        progress_placeholder.markdown(
            "\n".join(f"- {m}" for m in log_lines[-4:])
        )

    try:
        with st.spinner("Building template…"):
            parsed = parse_mapping_template(
                mapping_file.read(),
                filename=mapping_file.name,
            )
            result_bytes = build_output(
                blank_template_bytes=blank_bytes,
                parsed=parsed,
                source_systems=source_systems,
                progress=update_progress,
            )

        progress_placeholder.empty()

        object_name = parsed["meta"]["object_name"]
        field_count = len(parsed["meta"]["api_names_ordered"])

        st.success(f"✅ Done! Generated {field_count} field rows for **{object_name}**.")

        final_name = output_name or \
            f"UWM_UAT_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

        st.download_button(
            label=f"⬇️  Download  {final_name}",
            data=result_bytes,
            file_name=final_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        with st.expander("What was generated", expanded=True):
            branch_lines = []
            if "Andar" in source_systems:
                branch_lines.append("- WHALIF, WPEI (Andar branches)")
            if "DTracker" in source_systems:
                branch_lines.append("- WFREDE, WSTJOH (DTracker branches)")
            if "Raiser's Edge" in source_systems:
                branch_lines.append("- WRE1, WRE2 (Raiser's Edge branches)")

            st.markdown(f"""
**Object:** {object_name}
**Fields:** {field_count}
**Source systems:** {", ".join(source_systems)}

**Template sheet** includes:
- Andar Location, DT Location, Mapping Notes columns
- SCRM Field Name and SCRM Field Label populated from mapping template
- Calibri 12, thin borders throughout

**Branch sheets:**
{"".join(branch_lines) if branch_lines else "None"}
""")

    except ValueError as e:
        progress_placeholder.empty()
        st.error(f"❌ {e}")
        st.markdown(
            "**Tip:** Make sure you uploaded the strategy mapping template "
            "(the file with 'Template' or 'wstrat' in the name), not the UAT file."
        )
    except Exception as e:
        progress_placeholder.empty()
        st.error(f"❌ Unexpected error: {e}")
        st.markdown(
            "Please check the file is a valid `.xlsx` and try again. "
            "If the problem persists, contact your data team."
        )

# ── Footer ─────────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption(
    "UAT Template Generator · United Way Centraide Canada · "
    "For issues contact your data migration team."
)
