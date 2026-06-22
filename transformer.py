"""
transformer.py
Core transformation logic for UAT Excel files.
Reads a UAT tracking file + a strategy mapping template and produces
a fully formatted output Excel workbook.
"""

import re
from copy import copy
from io import BytesIO

import openpyxl
import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# ── Styling constants ──────────────────────────────────────────────────────────
FONT_NAME = "Calibri"
FONT_SIZE = 12
THIN = Side(style="thin")
THIN_BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def std_font(bold=False):
    return Font(name=FONT_NAME, size=FONT_SIZE, bold=bold)


def apply_std(cell, bold=False, wrap=False, border=True):
    cell.font = std_font(bold=bold)
    cell.alignment = Alignment(wrap_text=wrap)
    if border:
        cell.border = copy(THIN_BORDER)
    else:
        cell.border = Border()


def copy_style(src, dst):
    if src.has_style:
        dst.font = copy(src.font)
        dst.fill = copy(src.fill)
        dst.border = copy(src.border)
        dst.alignment = copy(src.alignment)
        dst.number_format = src.number_format


def swap_cols(ws, col_a, col_b, min_row=1):
    """Swap values and styles between two column indices across all rows."""
    for r in range(min_row, ws.max_row + 1):
        ca, cb = ws.cell(row=r, column=col_a), ws.cell(row=r, column=col_b)
        ca.value, cb.value = cb.value, ca.value
        for attr in ("font", "fill", "border", "alignment"):
            a_val, b_val = copy(getattr(ca, attr)), copy(getattr(cb, attr))
            setattr(ca, attr, b_val)
            setattr(cb, attr, a_val)


# ── Template mapping sheet parser ─────────────────────────────────────────────

# Keywords used to identify key rows (case-insensitive, partial match)
_ROW_KEYWORDS = {
    "api_name":    ["npsp api name", "api name", "api names"],
    "andar_field": ["andar field:", "andar field -", "andar api field", "andar field\n"],
    "andar_logic": ["andar logic", "logic\n", "^logic$"],
    "dt_field":    ["dtracker field", "dt field", "dt api field"],
    "dt_logic":    ["dtracker logic", "dt logic"],
}


def _match_keyword(label: str, keywords: list[str]) -> bool:
    label = label.strip().lower()
    for kw in keywords:
        if kw.startswith("^") and kw.endswith("$"):
            if re.fullmatch(kw[1:-1], label):
                return True
        elif kw in label:
            return True
    return False


def parse_mapping_template(template_bytes: bytes, source_systems: list[str]) -> dict:
    """
    Parse the strategy mapping template Excel file.
    Returns a dict keyed by SCRM API name with values:
        {
            "andar_field": str,
            "andar_logic": str,
            "dt_field":    str,
            "dt_logic":    str,
        }
    """
    # Find the mapping sheet (not Legend)
    wb = openpyxl.load_workbook(BytesIO(template_bytes), data_only=True)
    sheet_name = next(
        (s for s in wb.sheetnames if s.lower() not in ("legend",)), wb.sheetnames[0]
    )
    df = pd.read_excel(BytesIO(template_bytes), sheet_name=sheet_name, header=None)

    # Identify key rows by scanning col 0 labels
    row_map = {k: None for k in _ROW_KEYWORDS}
    for i, row in df.iterrows():
        label = str(row.iloc[0]).strip()
        if label in ("nan", ""):
            continue
        for key, keywords in _ROW_KEYWORDS.items():
            # Only assign the FIRST matching row for each key
            if row_map[key] is None and _match_keyword(label, keywords):
                row_map[key] = i
                break

    if row_map["api_name"] is None:
        raise ValueError(
            "Could not find an API Name row in the mapping template. "
            "Make sure the first data sheet has a row labelled 'NPSP API Name' or 'API Names'."
        )

    # Build col→API name index
    api_row = df.iloc[row_map["api_name"]].tolist()
    api_clean = [
        str(v).strip().split("\n")[0].strip() if str(v) != "nan" else ""
        for v in api_row
    ]

    def get_row_values(row_idx):
        if row_idx is None:
            return {}
        row = df.iloc[row_idx].tolist()
        return {
            api_clean[i]: str(v).strip()
            for i, v in enumerate(row)
            if api_clean[i] and str(v).strip() not in ("nan", "")
        }

    andar_fields = get_row_values(row_map["andar_field"])
    andar_logics = get_row_values(row_map["andar_logic"])
    dt_fields    = get_row_values(row_map["dt_field"])
    dt_logics    = get_row_values(row_map["dt_logic"])

    # Merge into per-field dicts
    all_apis = set(andar_fields) | set(andar_logics) | set(dt_fields) | set(dt_logics)
    result = {}
    for api in all_apis:
        result[api] = {
            "andar_field": andar_fields.get(api, ""),
            "andar_logic": andar_logics.get(api, ""),
            "dt_field":    dt_fields.get(api, ""),
            "dt_logic":    dt_logics.get(api, ""),
        }

    return result


# ── UAT file analyser ─────────────────────────────────────────────────────────

def detect_uat_layout(ws) -> dict:
    """
    Auto-detect layout of a UAT sheet:
    - header_row: row number of the column headers
    - data_start / data_end: row range of field data
    - meta_rows: list of (row_num, label) for tester/ID meta rows above the header
    - col positions for Andar, SCRM Field Name, SCRM Field Label, Mapping Notes
    - uat_col_start: first UAT data column (Payment 1, Relationship 1, etc.)
    """
    layout = {
        "header_row": None,
        "data_start": None,
        "data_end": None,
        "meta_rows": [],
        "andar_col": None,
        "scrm_name_col": None,
        "scrm_label_col": None,
        "mapping_col": None,
        "uat_col_start": None,
        "uat_col_header": None,
    }

    # Scan for header row: look for a row containing SCRM/Andar/Mapping in cols A-F
    for row in ws.iter_rows(min_row=1, max_row=20):
        vals = [str(c.value or "").strip().lower() for c in row[:8]]
        joined = " ".join(vals)
        if any(kw in joined for kw in ["scrm field", "andar field", "andar location", "mapping notes"]):
            layout["header_row"] = row[0].row
            # Identify each column by header value
            for cell in row:
                v = str(cell.value or "").strip().lower()
                if not v:
                    continue
                if "andar" in v and ("field" in v or "location" in v):
                    layout["andar_col"] = cell.column
                elif "scrm field name" in v or "scrm field name" in v:
                    layout["scrm_name_col"] = cell.column
                elif "scrm field label" in v:
                    layout["scrm_label_col"] = cell.column
                elif "mapping" in v:
                    layout["mapping_col"] = cell.column
                elif any(kw in v for kw in [
                    "payment", "relationship", "allocation", "gau",
                    "address", "affiliation", "federation"
                ]):
                    if layout["uat_col_start"] is None:
                        layout["uat_col_start"] = cell.column
                        layout["uat_col_header"] = str(cell.value or "").strip()
            break

    if layout["header_row"] is None:
        raise ValueError(
            "Could not detect the header row in the UAT file. "
            "Expected a row with 'Andar Field', 'SCRM Field Name', etc."
        )

    hr = layout["header_row"]

    # Meta rows = rows above header that have a label in col D (or wherever)
    for r in range(1, hr):
        for col_idx in range(1, 8):
            v = ws.cell(row=r, column=col_idx).value
            if v and str(v).strip():
                layout["meta_rows"].append((r, col_idx, str(v).strip()))
                break

    # Data rows: from header+1 until blank SCRM name col OR end of sheet
    scrm_col = layout["scrm_name_col"] or layout["andar_col"] or 1
    data_start = hr + 1
    data_end = data_start
    for r in range(data_start, ws.max_row + 1):
        # Check if ANY of cols A-E has a value (handle sparse rows)
        has_val = any(
            ws.cell(row=r, column=c).value not in (None, "")
            for c in range(1, 6)
        )
        if has_val:
            data_end = r
        else:
            # Allow one blank row gap, then stop
            next_has = any(
                ws.cell(row=r + 1, column=c).value not in (None, "")
                for c in range(1, 6)
            )
            if not next_has:
                break

    layout["data_start"] = data_start
    layout["data_end"] = data_end

    return layout


# ── Build combined mapping notes ───────────────────────────────────────────────

def build_notes(existing: str, api_name: str, field_data: dict, source_systems: list[str]) -> str:
    parts = []
    if existing and existing.strip():
        parts.append(existing.strip())

    fd = field_data.get(api_name, {})

    if "Andar" in source_systems and fd.get("andar_logic"):
        parts.append(f"Andar Logic: {fd['andar_logic']}")
    if "DTracker" in source_systems and fd.get("dt_logic"):
        parts.append(f"DTracker Logic: {fd['dt_logic']}")

    return "\n\n".join(parts) if parts else ""


# ── Main transformation ────────────────────────────────────────────────────────

BRANCH_CONFIGS = {
    "Andar":    [("WHALIF", "andar"), ("WPEI", "andar")],
    "DTracker": [("WFREDE", "dt"),    ("WSTJOH", "dt")],
}


def transform(
    uat_bytes: bytes,
    template_bytes: bytes,
    source_systems: list[str],
    progress_callback=None,
) -> bytes:
    """
    Full transformation pipeline.
    Returns the transformed Excel file as bytes.
    """

    def progress(msg):
        if progress_callback:
            progress_callback(msg)

    progress("Parsing mapping template…")
    field_data = parse_mapping_template(template_bytes, source_systems)

    progress("Loading UAT file…")
    wb = openpyxl.load_workbook(BytesIO(uat_bytes))

    # Rename Sheet1 / Sheet 1 → Template
    for name in list(wb.sheetnames):
        if name.strip().lower() in ("sheet1", "sheet 1"):
            wb[name].title = "Template"
            break

    # Process the Template sheet
    if "Template" not in wb.sheetnames:
        raise ValueError(
            "Could not find a sheet named 'Template', 'Sheet1', or 'Sheet 1' in the UAT file."
        )

    progress("Analysing UAT layout…")
    ws_t = wb["Template"]
    layout = detect_uat_layout(ws_t)
    hr = layout["header_row"]
    ds = layout["data_start"]
    de = layout["data_end"]

    # ── Determine which col is the SCRM API name (used for lookups) ───────────
    # Priority: explicit scrm_name_col, else andar_col (fallback for older files)
    scrm_lookup_col = layout["scrm_name_col"] or layout["andar_col"]

    # ── Rearrange columns to target order ─────────────────────────────────────
    # Target: A=Andar Location, B=DT Location(new), C=Mapping Notes,
    #         D=SCRM Field Name, E=SCRM Field Label, F+=UAT cols
    #
    # Step 1: ensure Andar col is col 1 (if it isn't already)
    if layout["andar_col"] and layout["andar_col"] != 1:
        swap_cols(ws_t, 1, layout["andar_col"])
        # Update layout col refs
        old_a = layout["andar_col"]
        for key in ("scrm_name_col", "scrm_label_col", "mapping_col", "uat_col_start"):
            if layout[key] == 1:
                layout[key] = old_a
        layout["andar_col"] = 1

    # Step 2: Move meta rows so they land above SCRM Field Label after insert
    # SCRM Field Label is currently at some col → after inserting col B it shifts +1
    # We want meta in the col that will become SCRM Field Label col after insert
    # i.e. currently (scrm_label_col - 1). Move them there.
    scrm_label_col = layout["scrm_label_col"]
    if scrm_label_col and scrm_label_col > 1:
        target_meta_col = scrm_label_col - 1  # will become scrm_label_col after insert
        for (r, cur_col, val) in layout["meta_rows"]:
            if cur_col != target_meta_col:
                ws_t.cell(row=r, column=target_meta_col).value = val
                ws_t.cell(row=r, column=cur_col).value = None

    # Step 3: Insert col B (DT Location)
    ws_t.insert_cols(2)
    # After insert all cols >= 2 shift +1
    for key in ("scrm_name_col", "scrm_label_col", "mapping_col", "uat_col_start"):
        if layout[key] and layout[key] >= 2:
            layout[key] += 1
    if scrm_lookup_col and scrm_lookup_col >= 2:
        scrm_lookup_col += 1

    # Step 4: Ensure Mapping Notes is col 3 (after Andar and DT Location)
    # If it's somewhere else, move it to col 3 by cascading swaps
    if layout["mapping_col"] and layout["mapping_col"] != 3:
        mc = layout["mapping_col"]
        # Shift left or right to col 3
        direction = -1 if mc > 3 else 1
        while mc != 3:
            swap_cols(ws_t, mc, mc + direction)
            mc += direction
        layout["mapping_col"] = 3
        # Recalculate scrm_name_col and scrm_label_col
        # After shuffling, re-read headers
        for cell in ws_t[hr]:
            v = str(cell.value or "").strip().lower()
            if "scrm field name" in v:
                layout["scrm_name_col"] = cell.column
                scrm_lookup_col = cell.column
            elif "scrm field label" in v:
                layout["scrm_label_col"] = cell.column

    # ── Update header row ──────────────────────────────────────────────────────
    progress("Updating headers…")
    ws_t.cell(row=hr, column=1).value = "Andar Location"
    ws_t.cell(row=hr, column=2).value = "DT Location"
    if not ws_t.cell(row=hr, column=3).value:
        ws_t.cell(row=hr, column=3).value = "Mapping Notes"

    # Copy fill from A header to B header
    ws_t.cell(row=hr, column=2).fill = copy(ws_t.cell(row=hr, column=1).fill)

    for col in range(1, ws_t.max_column + 1):
        c = ws_t.cell(row=hr, column=col)
        if c.value:
            apply_std(c, bold=True)

    # ── Style meta rows ────────────────────────────────────────────────────────
    new_meta_col = (scrm_label_col) if scrm_label_col else 4
    for r in range(1, hr):
        c = ws_t.cell(row=r, column=new_meta_col)
        if c.value:
            apply_std(c, bold=True)

    # ── Populate data rows ─────────────────────────────────────────────────────
    progress("Populating DT Location and Mapping Notes…")
    for row_num in range(ds, de + 1):
        # Get SCRM API name for lookup
        api_name = ""
        if scrm_lookup_col:
            api_name = str(ws_t.cell(row=row_num, column=scrm_lookup_col).value or "").strip()

        fd = field_data.get(api_name, {})

        # DT Location (col B = 2)
        dt_loc = fd.get("dt_field", "") if "DTracker" in source_systems else ""
        cell_b = ws_t.cell(row=row_num, column=2)
        cell_b.value = dt_loc if dt_loc else None

        # Mapping Notes (col 3)
        existing_note = str(ws_t.cell(row=row_num, column=3).value or "")
        combined = build_notes(existing_note, api_name, field_data, source_systems)
        cell_c = ws_t.cell(row=row_num, column=3)
        if combined:
            cell_c.value = combined

        # Andar Location (col A) — overwrite with andar_field if available
        if "Andar" in source_systems and fd.get("andar_field"):
            ws_t.cell(row=row_num, column=1).value = fd["andar_field"]

    # ── Apply Calibri 12 + borders to all data cells ───────────────────────────
    progress("Applying formatting…")
    for row in ws_t.iter_rows(min_row=ds, max_row=de):
        for cell in row:
            wrap = bool(cell.value and "\n" in str(cell.value))
            apply_std(cell, wrap=wrap)

    # ── DT Location border rule ────────────────────────────────────────────────
    # Border only if left (col A) or right (col C) is populated
    for r in range(ds, de + 1):
        left_val  = ws_t.cell(row=r, column=1).value
        right_val = ws_t.cell(row=r, column=3).value
        cell_b    = ws_t.cell(row=r, column=2)
        left_pop  = left_val not in (None, "", "nan")
        right_pop = right_val not in (None, "", "nan")
        if left_pop or right_pop:
            apply_std(cell_b, wrap=bool(cell_b.value and "\n" in str(cell_b.value or "")))
        else:
            cell_b.border    = Border()
            cell_b.font      = std_font()
            cell_b.alignment = Alignment()

    # ── Column widths ──────────────────────────────────────────────────────────
    ws_t.column_dimensions["A"].width = 40.0
    ws_t.column_dimensions["B"].width = 44.0
    ws_t.column_dimensions["C"].width = 44.0
    ws_t.column_dimensions["D"].width = 28.0
    ws_t.column_dimensions["E"].width = 28.0

    # ── Build branch sheets ────────────────────────────────────────────────────
    progress("Creating branch sheets…")

    def build_branch(sheet_name, branch_type):
        if sheet_name in wb.sheetnames:
            ws_b = wb[sheet_name]
            # Clear existing content and rebuild
            for row in ws_b.iter_rows():
                for cell in row:
                    cell.value = None
        else:
            ws_b = wb.create_sheet(sheet_name)

        # Meta rows
        for r in range(1, hr):
            c_src = ws_t.cell(row=r, column=new_meta_col)
            if c_src.value:
                apply_std(ws_b.cell(row=r, column=4, value=c_src.value), bold=True)

        # Header row
        loc_hdr = "DT Location" if branch_type == "dt" else "Andar Location"
        tmpl_fill = copy(ws_t.cell(row=hr, column=1).fill)

        # Detect UAT column header prefix ("Payment", "Relationship", etc.)
        uat_prefix = ""
        if layout["uat_col_start"]:
            raw = str(ws_t.cell(row=hr, column=layout["uat_col_start"]).value or "")
            # Strip trailing number
            uat_prefix = re.sub(r"\s*\d+$", "", raw).strip()

        branch_headers = [loc_hdr, "Mapping Notes", "SCRM Field Name", "SCRM Field Label"]
        uat_count = 25  # generous default
        uat_headers = [f"{uat_prefix} {n}" if uat_prefix else f"Record {n}" for n in range(1, uat_count + 1)]

        for col_i, hdr in enumerate(branch_headers + uat_headers, start=1):
            cell = ws_b.cell(row=hr, column=col_i, value=hdr)
            cell.fill = copy(tmpl_fill)
            apply_std(cell, bold=True)

        # Data rows
        for src_row in range(ds, de + 1):
            andar_v = ws_t.cell(row=src_row, column=1).value
            dt_v    = ws_t.cell(row=src_row, column=2).value
            note_v  = ws_t.cell(row=src_row, column=3).value
            name_v  = ws_t.cell(row=src_row, column=4).value if layout["scrm_name_col"] else None
            label_v = ws_t.cell(row=src_row, column=5).value if layout["scrm_label_col"] else None

            col_a_val = dt_v if branch_type == "dt" else andar_v
            row_data  = [col_a_val, note_v, name_v, label_v]

            for col_i, val in enumerate(row_data, start=1):
                cell = ws_b.cell(row=src_row, column=col_i, value=val)
                apply_std(cell, wrap=bool(val and "\n" in str(val)))

        # Column widths
        ws_b.column_dimensions["A"].width = 44.0
        ws_b.column_dimensions["B"].width = 44.0
        ws_b.column_dimensions["C"].width = 28.0
        ws_b.column_dimensions["D"].width = 28.0

    for system in source_systems:
        for sheet_name, btype in BRANCH_CONFIGS.get(system, []):
            build_branch(sheet_name, btype)

    # ── Also apply structural fixes to any existing tester sheets ─────────────
    progress("Updating existing tester sheets…")
    skip_sheets = {"Template", "WFREDE", "WSTJOH", "WHALIF", "WPEI", "Legend", "MAPPING", "Mapping"}
    for sname in wb.sheetnames:
        if sname in skip_sheets:
            continue
        ws_b = wb[sname]
        try:
            blayout = detect_uat_layout(ws_b)
            if blayout["header_row"] is None:
                continue
            bhr = blayout["header_row"]
            bds = blayout["data_start"]
            bde = blayout["data_end"]

            # Insert DT col if not present
            has_dt = any(
                "dt location" in str(ws_b.cell(row=bhr, column=c).value or "").lower()
                for c in range(1, 8)
            )
            if not has_dt:
                ws_b.insert_cols(2)
                ws_b.cell(row=bhr, column=2).value = "DT Location"
                ws_b.cell(row=bhr, column=2).fill = copy(ws_b.cell(row=bhr, column=1).fill)

            # Update headers and apply formatting
            ws_b.cell(row=bhr, column=1).value = "Andar Location"
            for col in range(1, ws_b.max_column + 1):
                c = ws_b.cell(row=bhr, column=col)
                if c.value:
                    apply_std(c, bold=True)

            # Copy DT Location and notes from Template
            for src_row in range(bds, bde + 1):
                dt_val   = ws_t.cell(row=src_row, column=2).value
                note_val = ws_t.cell(row=src_row, column=3).value
                ws_b.cell(row=src_row, column=2).value = dt_val
                ws_b.cell(row=src_row, column=3).value = note_val
                apply_std(ws_b.cell(row=src_row, column=2),
                          wrap=bool(dt_val and "\n" in str(dt_val or "")))
                apply_std(ws_b.cell(row=src_row, column=3),
                          wrap=bool(note_val and "\n" in str(note_val or "")))

            # Apply formatting to all data cells in tester sheet
            for row in ws_b.iter_rows(min_row=bds, max_row=bde):
                for cell in row:
                    if cell.column == 2:
                        continue  # already handled
                    if cell.value is not None:
                        apply_std(cell, wrap=bool(cell.value and "\n" in str(cell.value)))

        except Exception:
            pass  # Don't fail on tester sheets — they're secondary

    progress("Saving…")
    out = BytesIO()
    wb.save(out)
    return out.getvalue()
