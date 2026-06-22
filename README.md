# UAT Template Transformer

A no-code web tool that transforms UAT tracking Excel files for Salesforce/NPSP data migrations.

## What it does

Upload two Excel files → click Generate → download your transformed file.

**Specifically it:**
- Renames the main sheet to `Template`
- Inserts an `Andar Location` column, `DT Location` column, and `Mapping Notes` column in the correct order
- Populates those columns from the strategy mapping template (matched by SCRM API name)
- Creates branch sheets (WHALIF/WPEI for Andar, WFREDE/WSTJOH for DTracker)
- Applies Calibri 12 font and thin borders throughout
- Updates any existing tester sheets (e.g. Ally, Siddig) with the new structure

## Deploy to Streamlit Community Cloud (free, ~5 minutes)

### Prerequisites
- A [GitHub account](https://github.com)
- A [Streamlit Community Cloud account](https://streamlit.io/cloud) (free, sign in with GitHub)

### Steps

1. **Create a new GitHub repository**
   - Go to github.com → New repository
   - Name it `uat-transformer` (or anything you like)
   - Set it to **Public** (required for free Streamlit hosting)
   - Click Create repository

2. **Upload these files to the repository**
   Upload all three files:
   - `app.py`
   - `transformer.py`
   - `requirements.txt`

   You can drag-and-drop them in the GitHub web interface.

3. **Deploy on Streamlit**
   - Go to [share.streamlit.io](https://share.streamlit.io)
   - Click **New app**
   - Select your GitHub repo and branch (`main`)
   - Set **Main file path** to `app.py`
   - Click **Deploy**

4. **Share the link**
   Streamlit gives you a URL like:
   `https://your-name-uat-transformer-app-xyz123.streamlit.app`

   Share this with your team — no login, no install, just a URL.

## Making it private (optional)

If you don't want the app publicly accessible:
- In Streamlit Cloud → Settings → Sharing → set to **Only specific people**
- Add team members by email

Or: keep the GitHub repo public but the app viewer-restricted.

## Local development (optional)

If you want to run it on your own machine:

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Adding new objects

The transformer is designed to work generically across all UAT objects (GAU, Payment, Allocation, Address, Affiliation, Account Relationships, Individual Relationships, and future ones).

It detects the layout automatically by scanning for known header keywords. If a new object's template uses unusual row labels, update the `_ROW_KEYWORDS` dictionary in `transformer.py`.

## Transferring to Microsoft (Power Automate / Teams)

The core logic lives in `transformer.py` — pure Python with `openpyxl` and `pandas`.
This same file can be deployed as an **Azure Function** and triggered from Power Automate,
making it accessible as a button inside Microsoft Teams with no changes to the logic.
