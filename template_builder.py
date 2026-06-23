"""
template_builder.py
Builds the output Excel workbook from scratch using:
  - The hardcoded blank UAT template
  - Parsed mapping data from mapping_parser.py
  - User's source system selections and custom branch names

Dynamic column layout: unused DB columns are omitted when fewer
than 3 source systems are selected.
"""

from copy import copy
from io import BytesIO

import openpyxl
from openpyxl.styles import Alignment, Border

from styles import apply_std, std_font, THIN_BORDER, NO_BORDER

HEADER_ROW = 5
DATA_START = 6


def _build_notes(api_name: str, fields: dict, source_systems: list) -> str:
    """Combine Andar Logic and DTracker Logic into Mapping Notes."""
    fd = fields.get(api_name, {})
    parts = []
    if "Andar" in source_systems and fd.get("andar_logic"):
        parts.append(f"Andar Logic: {fd['andar_logic']}")
    if "DTracker" in source_systems and fd.get("dt_logic"):
        parts.append(f"DTracker Logic: {fd['dt_logic']}")
    return "\n\n".join(parts)


def _col_layout(source_systems: list) -> dict:
    """
    Build column layout dynamically based on selected source systems.
    Only includes DB columns for selected systems — no empty placeholders.

    Returns dict: {col_name: col_number (1-based)}
    """
    layout = {}
    col = 1

    if "Andar" in source_systems:
        layout["andar"] = col; col += 1
    if "DTracker" in source_systems:
        layout["dt"] = col; col += 1
    if "Raiser's Edge" in source_systems:
        layout["re"] = col; col += 1

    layout["notes"]    = col; col += 1
    layout["api_name"] = col; col += 1
    layout["label"]    = col; col += 1
    layout["uat_start"] = col

    return layout


def build_output(
    blank_template_bytes: bytes,
    parsed: dict,
    source_systems: list,
    branch_overrides: list = None,
    progress=None,
) -> bytes:
    """
    Build the complete output workbook from the blank template.

    Args:
        blank_template_bytes : hardcoded blank template as bytes
        parsed               : result of mapping_parser.parse_mapping_template()
        source_systems       : e.g. ["Andar", "DTracker"]
        branch_overrides     : list of (sheet_name, branch_type) — custom names
                               e.g. [("WHALIF", "andar"), ("WFREDE", "dt")]
                               If None, uses defaults.
        progress             : optional callable(str) for status updates
    """
    def log(msg):
        if progress: progress(msg)

    fields            = parsed["fields"]
    meta              = parsed["meta"]
    object_name       = meta["object_name"]
    api_names_ordered = meta["api_names_ordered"]
    cols              = _col_layout(source_systems)
    meta_col          = cols["label"]   # meta rows sit above SCRM Field Label

    # Default branch names if none provided
    if branch_overrides is None:
        branch_overrides = _default_branches(source_systems)

    log("Loading blank template…")
    wb = openpyxl.load_workbook(BytesIO(blank_template_bytes))
    ws = wb["Template"]

    # Capture header fill before overwriting
    header_fill = copy(ws.cell(row=HEADER_ROW, column=1).fill)

    # Clear the entire template sheet first (it has placeholder content)
    for row in ws.iter_rows():
        for cell in row:
            cell.value = None

    # ── Meta rows ─────────────────────────────────────────────────────────────
    log("Setting meta rows…")
    meta_labels = ["Tester", f"UAT {object_name}",
                   "Original ID", f"UAT {object_name} ID"]
    for i, label in enumerate(meta_labels, start=1):
        apply_std(ws.cell(row=i, column=meta_col, value=label), bold=True)

    # ── Header row ─────────────────────────────────────────────────────────────
    log("Writing column headers…")
    db_header_map = {
        "andar": "Andar Location",
        "dt":    "DT Location",
        "re":    "RE Location",
    }
    fixed_header_map = {
        "notes":    "Mapping Notes",
        "api_name": "SCRM Field Name",
        "label":    "SCRM Field Label",
    }

    for key, col_num in cols.items():
        if key == "uat_start":
            continue
        label = db_header_map.get(key) or fixed_header_map.get(key, "")
        cell = ws.cell(row=HEADER_ROW, column=col_num, value=label)
        cell.fill = copy(header_fill)
        apply_std(cell, bold=True)

    # UAT data columns
    for n in range(1, 13):
        col_num = cols["uat_start"] + n - 1
        cell = ws.cell(row=HEADER_ROW, column=col_num,
                       value=f"{object_name} {n}")
        cell.fill = copy(header_fill)
        apply_std(cell, bold=True)

    # ── Data rows ──────────────────────────────────────────────────────────────
    log(f"Writing {len(api_names_ordered)} field rows…")
    for row_offset, api_name in enumerate(api_names_ordered):
        row_num = DATA_START + row_offset
        fd = fields.get(api_name, {})

        # DB location columns
        if "andar" in cols:
            val = fd.get("andar_field", "")
            apply_std(ws.cell(row=row_num, column=cols["andar"],
                              value=val or None),
                      wrap=bool(val and "\n" in val))

        if "dt" in cols:
            val = fd.get("dt_field", "")
            left_pop  = bool(fd.get("andar_field") if "andar" in cols else False)
            right_pop = bool(fd.get("andar_logic") or fd.get("dt_logic"))
            apply_std(ws.cell(row=row_num, column=cols["dt"],
                              value=val or None),
                      wrap=bool(val and "\n" in val),
                      border=(left_pop or right_pop or bool(val)))

        if "re" in cols:
            apply_std(ws.cell(row=row_num, column=cols["re"]))

        # Mapping Notes
        notes = _build_notes(api_name, fields, source_systems)
        apply_std(ws.cell(row=row_num, column=cols["notes"],
                          value=notes or None),
                  wrap=bool(notes and "\n" in notes))

        # SCRM Field Name
        apply_std(ws.cell(row=row_num, column=cols["api_name"],
                          value=api_name))

        # SCRM Field Label
        label = fd.get("scrm_label", "")
        apply_std(ws.cell(row=row_num, column=cols["label"],
                          value=label or None))

        # UAT columns
        for n in range(1, 13):
            apply_std(ws.cell(row=row_num,
                              column=cols["uat_start"] + n - 1))

    # ── Column widths ──────────────────────────────────────────────────────────
    width_map = {
        "andar": ("A", 40.0), "dt": ("B", 44.0), "re": ("C", 40.0),
        "notes": (None, 44.0), "api_name": (None, 30.0), "label": (None, 26.0),
    }
    from openpyxl.utils import get_column_letter
    for key, col_num in cols.items():
        if key == "uat_start":
            continue
        width = {
            "andar": 40.0, "dt": 44.0, "re": 40.0,
            "notes": 44.0, "api_name": 30.0, "label": 26.0,
        }.get(key, 18.0)
        ws.column_dimensions[get_column_letter(col_num)].width = width

    # ── Branch sheets ──────────────────────────────────────────────────────────
    log("Creating branch sheets…")
    for sheet_name, branch_type in branch_overrides:
        if not sheet_name:
            continue
        _build_branch(wb, sheet_name, branch_type,
                      object_name, api_names_ordered, fields,
                      source_systems, header_fill, meta_col)

    log("Saving…")
    out = BytesIO()
    wb.save(out)
    return out.getvalue()


def _default_branches(source_systems: list) -> list:
    """Return default (sheet_name, branch_type) pairs for selected systems."""
    defaults = {
        "Andar":         [("WHALIF", "andar")],
        "DTracker":      [("WFREDE", "dt")],
        "Raiser's Edge": [("WRE1",   "re")],
    }
    result = []
    for system in source_systems:
        result.extend(defaults.get(system, []))
    return result


def _build_branch(wb, sheet_name, branch_type,
                  object_name, api_names_ordered, fields,
                  source_systems, header_fill, meta_col):
    """Create a single branch sheet."""
    from openpyxl.utils import get_column_letter

    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)

    # Meta rows
    meta_labels = ["Tester", f"UAT {object_name}",
                   "Original ID", f"UAT {object_name} ID"]
    for i, label in enumerate(meta_labels, start=1):
        apply_std(ws.cell(row=i, column=meta_col, value=label), bold=True)

    # Header
    loc_hdr = {"andar": "Andar Location",
                "dt":    "DT Location",
                "re":    "RE Location"}.get(branch_type, "Location")

    headers = [loc_hdr, "Mapping Notes", "SCRM Field Name", "SCRM Field Label"]
    uat_headers = [f"{object_name} {n}" for n in range(1, 26)]

    for col_i, hdr in enumerate(headers + uat_headers, start=1):
        cell = ws.cell(row=HEADER_ROW, column=col_i, value=hdr)
        cell.fill = copy(header_fill)
        apply_std(cell, bold=True)

    # Data rows
    for row_offset, api_name in enumerate(api_names_ordered):
        row_num = DATA_START + row_offset
        fd = fields.get(api_name, {})

        loc_val = {
            "andar": fd.get("andar_field", ""),
            "dt":    fd.get("dt_field", ""),
            "re":    "",
        }.get(branch_type, "")

        notes = _build_notes(api_name, fields, source_systems)
        label = fd.get("scrm_label", "")

        for col_i, val in enumerate(
            [loc_val or None, notes or None, api_name, label or None], start=1
        ):
            apply_std(ws.cell(row=row_num, column=col_i, value=val),
                      wrap=bool(val and "\n" in str(val)))

    # Column widths
    for col_i, width in enumerate([44.0, 44.0, 30.0, 26.0], start=1):
        ws.column_dimensions[get_column_letter(col_i)].width = width
