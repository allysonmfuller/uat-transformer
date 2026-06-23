# UAT Template Generator v2

Generates a fully formatted UAT tracking Excel file from a single upload —
the strategy mapping template. No UAT tracking file needed.

## Files

| File | Purpose |
|------|---------|
| `app.py` | Streamlit UI — what the user sees |
| `mapping_parser.py` | Reads the strategy mapping template, extracts field data |
| `template_builder.py` | Builds the output Excel from the blank template + parsed data |
| `styles.py` | All font/border/formatting helpers — one place to change styling |
| `UWC_NPSP_User_Acceptance_Testing_Template.xlsx` | Hardcoded blank template |
| `requirements.txt` | Python package dependencies |

## How to deploy (Streamlit Community Cloud)

1. Push all 6 files to your GitHub repo (`uat-transformer`)
2. Go to share.streamlit.io → New app → select repo → main file: `app.py`
3. Share the URL

## How to run locally

```bash
venv\Scripts\activate
streamlit run app.py
```

## Adding a new object

No code changes needed — just upload its mapping template.
The app detects the object name from the filename and reads
whatever fields are in the MAPPING sheet automatically.

## Adding Raiser's Edge support

When you have a RE mapping template, add its row label keywords
to `ROW_KEYWORDS` in `mapping_parser.py` under `"re_field"` and `"re_logic"`,
then handle them in `template_builder.py` the same way as Andar/DTracker.
