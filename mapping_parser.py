"""
mapping_parser.py
Reads the strategy mapping template Excel file and extracts
per-field data (Andar location/logic, DTracker location/logic,
SCRM label) keyed by SCRM API name.
"""

import re
from io import BytesIO

import openpyxl
import pandas as pd


# ── Keyword lists for detecting each row in the mapping sheet ──────────────────
# Each key maps to a list of substrings (lowercase) that identify that row.
# The first matching row found wins.
ROW_KEYWORDS = {
    "api_name":    ["npsp api name", "api name", "api names"],
    "scrm_label":  ["andar label", "andar ui label"],
    "andar_field": ["andar field:", "andar api field", "andar field -",
                    "andar field - individuals", "andar field\n",
                    "andar field (import", "import header"],
    "andar_logic": ["andar logic", "^logic$"],
    "dt_field":    ["dtracker field", "dt field", "dt api field",
                    "dtracker field:", "dt field - individuals"],
    "dt_logic":    ["dtracker logic", "dt logic"],
    "object_name": ["object name", "salesforce object", "npsp object"],
}


def _match(label: str, keywords: list) -> bool:
    """Return True if label matches any keyword (case-insensitive)."""
    label = label.strip().lower()
    for kw in keywords:
        if kw.startswith("^") and kw.endswith("$"):
            if re.fullmatch(kw[1:-1], label):
                return True
        elif kw in label:
            return True
    return False


def _find_mapping_sheet(wb) -> str:
    """Return the name of the mapping sheet (not Legend)."""
    for name in wb.sheetnames:
        if name.lower() not in ("legend",):
            return name
    return wb.sheetnames[0]


def _infer_object_name(filename: str) -> str:
    """
    Try to extract a human-readable object name from the template filename.
    e.g. ObjS02_02_wstrat_npe01__OppPayment__c_Template.xlsx → Payment

    To add a new object: add a key/value pair to the `known` dict below.
    The key is any lowercase substring that appears in the filename.
    Order matters — more specific keys should come before general ones.
    """
    # Key = lowercase substring to match anywhere in the filename
    # Value = human-readable object name used in the output file
    known = {
        # ── Object Set 1-3 ────────────────────────────────────────────────────
        "opppayment":              "Payment",
        "oppayment":               "Payment",
        "payment":                 "Payment",
        "allocation":              "Allocation",
        "address":                 "Address",
        "affiliation":             "Affiliation",
        "accountrelationship":     "Account Relationship",
        "individualrelationship":  "Individual Relationship",
        "generalaccountingunit":   "GAU",
        "gau":                     "GAU",
        "contact":                 "Contact",
        "opportunity":             "Opportunity",
        "account":                 "Account",
        # ── Object Set 4 ─────────────────────────────────────────────────────
        "campaignmember":          "Campaign Member",
        "campaign_member":         "Campaign Member",
        "volunteer_job":           "Volunteer Job",
        "volunteerjob":            "Volunteer Job",
        "teams__c":                "Teams",
        "volunteerteammember":     "Volunteer Team Member",
        "volunteer_hours":         "Volunteer Hours",
        "volunteerhours":          "Volunteer Hours",
        "note_and_communication":  "Note and Communication Code",
        "noteandcommunication":    "Note and Communication Code",
        "communication_log":       "Communication Log Task",
        "communicationlog":        "Communication Log Task",
        "attachment":              "Attachment",
        "notes_task":              "Notes Task",
        "notestask":               "Notes Task",
        "engagement_plan_task":    "Engagement Plan Task",
        "engagementplantask":      "Engagement Plan Task",
        "engagement_plan__c":      "Engagement Plan",
        "engagementplan":          "Engagement Plan",
        # ── Generic fallbacks — must stay at the bottom ───────────────────────
        "campaign":                "Campaign",
    }

    lower = filename.lower()
    for key, label in known.items():
        if key in lower:
            return label

    # Fallback: grab the part after last __ before _Template and split CamelCase
    match = re.search(r"__([A-Za-z]+)__c", filename)
    if match:
        raw = match.group(1)
        spaced = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", raw)
        return spaced.strip()

    return "Object"


def parse_mapping_template(template_bytes: bytes, filename: str = "") -> dict:
    """
    Parse the strategy mapping template.

    Returns a dict with two keys:
        "fields"  : dict keyed by SCRM API name ->
                        {scrm_label, andar_field, andar_logic, dt_field, dt_logic}
        "meta"    : dict with object-level info ->
                        {object_name, api_names_ordered (ordered list)}
    """
    wb         = openpyxl.load_workbook(BytesIO(template_bytes), data_only=True)
    sheet_name = _find_mapping_sheet(wb)
    df         = pd.read_excel(BytesIO(template_bytes), sheet_name=sheet_name, header=None)

    # ── Identify key rows by scanning col 0 ───────────────────────────────────
    row_map = {k: None for k in ROW_KEYWORDS}
    for i, row in df.iterrows():
        label = str(row.iloc[0]).strip()
        if label in ("nan", ""):
            continue
        for key, keywords in ROW_KEYWORDS.items():
            if row_map[key] is None and _match(label, keywords):
                row_map[key] = i
                break

    if row_map["api_name"] is None:
        raise ValueError(
            "Could not find an API Name row in the mapping template. "
            "Expected a row labelled 'NPSP API Name' or 'API Names'."
        )

    # ── Build col → API name index ─────────────────────────────────────────────
    api_row   = df.iloc[row_map["api_name"]].tolist()
    api_clean = []
    for v in api_row:
        s = str(v).strip().split("\n")[0].strip()
        api_clean.append(s if s != "nan" else "")

    # Skip col 0 (it's the row-label column) and collect all field API names
    api_names_ordered = [a for a in api_clean[1:] if a]

    def row_to_dict(row_idx):
        """Turn a mapping row into a {api_name: value} dict."""
        if row_idx is None:
            return {}
        row = df.iloc[row_idx].tolist()
        return {
            api_clean[i]: str(v).strip()
            for i, v in enumerate(row)
            if i > 0 and api_clean[i] and str(v).strip() not in ("nan", "")
        }

    scrm_labels  = row_to_dict(row_map["scrm_label"])
    andar_fields = row_to_dict(row_map["andar_field"])
    andar_logics = row_to_dict(row_map["andar_logic"])
    dt_fields    = row_to_dict(row_map["dt_field"])
    dt_logics    = row_to_dict(row_map["dt_logic"])

    # ── Merge into per-field dicts ─────────────────────────────────────────────
    all_apis = set(api_names_ordered)
    fields   = {}
    for api in all_apis:
        fields[api] = {
            "scrm_label":  scrm_labels.get(api, ""),
            "andar_field": andar_fields.get(api, ""),
            "andar_logic": andar_logics.get(api, ""),
            "dt_field":    dt_fields.get(api, ""),
            "dt_logic":    dt_logics.get(api, ""),
        }

    object_name = _infer_object_name(filename)

    return {
        "fields": fields,
        "meta": {
            "object_name":       object_name,
            "api_names_ordered": api_names_ordered,
        },
    }
