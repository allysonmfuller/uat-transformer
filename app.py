"""
app.py
Streamlit UI for the UAT Template Generator.

Two tabs:
  Tab 1 - Update Existing: upload UAT file + mapping template → enriched output
  Tab 2 - Generate from Scratch: upload mapping template only → new UAT file
"""

import re
from datetime import datetime
from pathlib import Path

import streamlit as st

from mapping_parser import parse_mapping_template
from template_builder import build_output
from transformer import transform

# ── Page config ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="UAT Template Generator",
    page_icon="📊",
    layout="centered",
)

st.markdown("""
<style>
    .block-container { padding-top: 2rem; max-width: 820px; }
    .step-label {
        font-size: 0.78rem;
        font-weight: 600;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        color: #888;
        margin-bottom: 0.25rem;
    }
    .divider { margin: 1.4rem 0 1rem; border-top: 1px solid #eee; }
    .branch-row {
        display: flex;
        align-items: center;
        gap: 8px;
        margin-bottom: 6px;
    }
</style>
""", unsafe_allow_html=True)

# ── Load blank template ────────────────────────────────────────────────────────
BLANK_PATH = Path(__file__).parent / "UWC_NPSP_User_Acceptance_Testing_Template.xlsx"

@st.cache_data
def load_blank() -> bytes:
    with open(BLANK_PATH, "rb") as f:
        return f.read()

try:
    blank_bytes = load_blank()
except FileNotFoundError:
    st.error(
        "❌ Blank template file not found. "
        "Make sure `UWC_NPSP_User_Acceptance_Testing_Template.xlsx` "
        "is in the same folder as `app.py`."
    )
    st.stop()

# ── Header ─────────────────────────────────────────────────────────────────────
st.title("📊 UAT Template Generator")
st.caption("United Way Centraide Canada · Salesforce/NPSP Data Migration")

# ── Tabs ───────────────────────────────────────────────────────────────────────
tab1, tab2 = st.tabs(["✏️  Update Existing UAT File", "🆕  Generate from Scratch"])


# ══════════════════════════════════════════════════════════════════════════════
# Shared helper: source system + branch name UI
# Returns: (source_systems list, branches dict {system: [sheet_name, ...]})
# ══════════════════════════════════════════════════════════════════════════════
def source_system_ui(key_prefix: str):
    """
    Renders source system checkboxes and dynamic branch name inputs.
    key_prefix keeps Streamlit widget keys unique between tabs.
    Returns (source_systems, branches).
    """
    st.markdown("Which source system(s) is this migration coming **from**?")

    col1, col2, col3 = st.columns(3)
    with col1:
        use_andar = st.checkbox("Andar",          value=True,  key=f"{key_prefix}_andar")
    with col2:
        use_dt    = st.checkbox("DTracker",       value=True,  key=f"{key_prefix}_dt")
    with col3:
        use_re    = st.checkbox("Raiser's Edge",  value=False, key=f"{key_prefix}_re")

    source_systems = []
    if use_andar: source_systems.append("Andar")
    if use_dt:    source_systems.append("DTracker")
    if use_re:    source_systems.append("Raiser's Edge")

    if not source_systems:
        st.warning("Please select at least one source system.")

    # ── Branch naming ──────────────────────────────────────────────────────────
    DEFAULTS = {
        "Andar":         ("WHALIF", "andar"),
        "DTracker":      ("WFREDE", "dt"),
        "Raiser's Edge": ("WRE1",   "re"),
    }

    branches = {}   # {system: [(sheet_name, branch_type), ...]}

    if source_systems:
        st.markdown('<div class="divider"></div>', unsafe_allow_html=True)
        st.markdown('<p class="step-label">Branch Sheet Names</p>',
                    unsafe_allow_html=True)
        st.caption(
            "Each source system gets one branch sheet by default. "
            "Click **+ Add branch** to add more for that system."
        )

        for system in source_systems:
            default_name, btype = DEFAULTS[system]

            # Count key tracks how many branches this system has
            count_key = f"{key_prefix}_{system}_count"
            if count_key not in st.session_state:
                st.session_state[count_key] = 1

            st.markdown(f"**{system}**")

            branch_names = []
            for i in range(st.session_state[count_key]):
                col_input, col_remove = st.columns([4, 1])
                with col_input:
                    default_val = default_name if i == 0 else ""
                    name = st.text_input(
                        f"Branch {i+1}",
                        value=st.session_state.get(
                            f"{key_prefix}_{system}_branch_{i}", default_val
                        ),
                        placeholder=f"e.g. {default_name}",
                        key=f"{key_prefix}_{system}_branch_{i}",
                        label_visibility="collapsed",
                    )
                    branch_names.append((name.strip(), btype))
                with col_remove:
                    if i > 0:  # Can't remove the first branch
                        if st.button("✕", key=f"{key_prefix}_{system}_remove_{i}",
                                     help="Remove this branch"):
                            st.session_state[count_key] -= 1
                            st.rerun()

            if st.button(f"+ Add branch", key=f"{key_prefix}_{system}_add",
                         help=f"Add another {system} branch sheet"):
                st.session_state[count_key] += 1
                st.rerun()

            # Filter out any blank names
            branches[system] = [(n, t) for n, t in branch_names if n]

    return source_systems, branches


def suggested_output_name(mapping_file, uat_file=None):
    """Auto-suggest output filename from uploaded file names."""
    base = mapping_file.name if mapping_file else (uat_file.name if uat_file else "")
    raw = base.replace(".xlsx", "")
    raw = re.sub(r"_w[a-z]+_", "_", raw)
    raw = re.sub(r"_Template$", "", raw, flags=re.IGNORECASE)
    raw = re.sub(r"(?i)uwem", "UWM", raw)
    raw = re.sub(r"_temp$", "", raw, flags=re.IGNORECASE)
    return f"UWM_{raw}_UAT_Round_1.xlsx"


def progress_logger():
    """Returns a placeholder and a logging callback."""
    placeholder = st.empty()
    lines = []
    def log(msg):
        lines.append(msg)
        placeholder.markdown("\n".join(f"- {m}" for m in lines[-4:]))
    return placeholder, log


def show_summary(parsed, source_systems, branches):
    """Show a summary expander after successful generation."""
    object_name = parsed["meta"]["object_name"]
    field_count = len(parsed["meta"]["api_names_ordered"])

    st.success(f"✅ Done! Generated {field_count} field rows for **{object_name}**.")

    branch_lines = []
    for system, branch_list in branches.items():
        names = ", ".join(n for n, _ in branch_list if n)
        if names:
            branch_lines.append(f"- {system}: {names}")

    with st.expander("What was generated", expanded=True):
        st.markdown(f"""
**Object:** {object_name}
**Fields:** {field_count}
**Source systems:** {", ".join(source_systems)}

**Branch sheets:**
{"".join(branch_lines) if branch_lines else "None"}

**Formatting:** Calibri 12, thin borders, Mapping Notes populated from template
""")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — Update Existing UAT File
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.markdown("#### Update an existing UAT tracking file")
    st.caption(
        "Use this when you already have a UAT file started. "
        "The app will add DT Location, enrich Mapping Notes, and create branch sheets."
    )
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # Step 1: Source systems
    st.markdown('<p class="step-label">Step 1 — Source Systems & Branch Names</p>',
                unsafe_allow_html=True)
    t1_systems, t1_branches = source_system_ui("t1")

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # Step 2: File uploads
    st.markdown('<p class="step-label">Step 2 — Upload Files</p>',
                unsafe_allow_html=True)

    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("**Existing UAT Tracking File**")
        st.caption("e.g. `GAU_-_UWEM_Data_Migration_UAT_-_Round_1.xlsx`")
        t1_uat_file = st.file_uploader(
            "UAT file", type=["xlsx"], label_visibility="collapsed", key="t1_uat"
        )
    with col_b:
        st.markdown("**Strategy Mapping Template**")
        st.caption("e.g. `ObjS01_05_wmarit_..._Template.xlsx`")
        t1_map_file = st.file_uploader(
            "Mapping template", type=["xlsx"], label_visibility="collapsed", key="t1_map"
        )

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # Step 3: Output name
    st.markdown('<p class="step-label">Step 3 — Output File Name (optional)</p>',
                unsafe_allow_html=True)
    t1_suggested = suggested_output_name(t1_map_file, t1_uat_file) \
        if (t1_map_file or t1_uat_file) else ""
    t1_output = st.text_input(
        "Output filename", value=t1_suggested,
        placeholder="UWM_GAU_UAT_Round_1.xlsx",
        label_visibility="collapsed", key="t1_outname"
    )
    if t1_output and not t1_output.endswith(".xlsx"):
        t1_output += ".xlsx"

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # Step 4: Generate
    st.markdown('<p class="step-label">Step 4 — Generate</p>', unsafe_allow_html=True)
    t1_ready = t1_uat_file and t1_map_file and len(t1_systems) > 0

    if st.button("⚙️  Update UAT File", disabled=not t1_ready,
                 use_container_width=True, type="primary", key="t1_btn"):
        ph, log = progress_logger()
        try:
            with st.spinner("Updating…"):
                parsed = parse_mapping_template(
                    t1_map_file.read(), filename=t1_map_file.name
                )
                # Flatten branches dict → list of (sheet_name, branch_type)
                flat_branches = [
                    (name, btype)
                    for system, pairs in t1_branches.items()
                    for name, btype in pairs
                ]
                result = transform(
                    uat_bytes=t1_uat_file.read(),
                    template_bytes=t1_map_file.read() if hasattr(t1_map_file, 'read') else t1_map_file,
                    source_systems=t1_systems,
                    branch_overrides=flat_branches,
                    progress_callback=log,
                )
            ph.empty()
            final = t1_output or f"UWM_UAT_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            st.download_button(
                f"⬇️  Download  {final}", data=result, file_name=final,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            show_summary(parsed, t1_systems, t1_branches)
        except ValueError as e:
            ph.empty()
            st.error(f"❌ {e}")
        except Exception as e:
            ph.empty()
            st.error(f"❌ Unexpected error: {e}")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — Generate from Scratch
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.markdown("#### Generate a new UAT file from a mapping template")
    st.caption(
        "Use this for new objects with no UAT file yet. "
        "Upload the strategy mapping template and the app builds everything."
    )
    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # Step 1: Source systems
    st.markdown('<p class="step-label">Step 1 — Source Systems & Branch Names</p>',
                unsafe_allow_html=True)
    t2_systems, t2_branches = source_system_ui("t2")

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # Step 2: Upload mapping template
    st.markdown('<p class="step-label">Step 2 — Upload Strategy Mapping Template</p>',
                unsafe_allow_html=True)
    st.caption("e.g. `ObjS04_06_wstrat_Volunteers__Volunteer_Job__c_Template.xlsx`")
    t2_map_file = st.file_uploader(
        "Mapping template", type=["xlsx"], label_visibility="collapsed", key="t2_map"
    )

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # Step 3: Output name
    st.markdown('<p class="step-label">Step 3 — Output File Name (optional)</p>',
                unsafe_allow_html=True)
    t2_suggested = suggested_output_name(t2_map_file) if t2_map_file else ""
    t2_output = st.text_input(
        "Output filename", value=t2_suggested,
        placeholder="UWM_VolunteerJob_UAT_Round_1.xlsx",
        label_visibility="collapsed", key="t2_outname"
    )
    if t2_output and not t2_output.endswith(".xlsx"):
        t2_output += ".xlsx"

    st.markdown('<div class="divider"></div>', unsafe_allow_html=True)

    # Step 4: Generate
    st.markdown('<p class="step-label">Step 4 — Generate</p>', unsafe_allow_html=True)
    t2_ready = t2_map_file is not None and len(t2_systems) > 0

    if st.button("⚙️  Generate UAT Template", disabled=not t2_ready,
                 use_container_width=True, type="primary", key="t2_btn"):
        ph, log = progress_logger()
        try:
            with st.spinner("Building…"):
                parsed = parse_mapping_template(
                    t2_map_file.read(), filename=t2_map_file.name
                )
                # Flatten branches dict → list of (sheet_name, branch_type)
                flat_branches = [
                    (name, btype)
                    for system, pairs in t2_branches.items()
                    for name, btype in pairs
                ]
                result = build_output(
                    blank_template_bytes=blank_bytes,
                    parsed=parsed,
                    source_systems=t2_systems,
                    branch_overrides=flat_branches,
                    progress=log,
                )
            ph.empty()
            final = t2_output or f"UWM_UAT_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
            st.download_button(
                f"⬇️  Download  {final}", data=result, file_name=final,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
            show_summary(parsed, t2_systems, t2_branches)
        except ValueError as e:
            ph.empty()
            st.error(f"❌ {e}")
        except Exception as e:
            ph.empty()
            st.error(f"❌ Unexpected error: {e}")

# ── Footer ─────────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption(
    "UAT Template Generator · United Way Centraide Canada · "
    "For issues contact your data migration team."
)
