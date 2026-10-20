# EWB Operations Portal

A unified Streamlit web app for **SP MADRID EastWest Bank Recovery Operations** — three tools in one portal.

## Tools Inside

| Tab | Tool | Purpose |
|-----|------|---------|
| 📑 | **Endorsement Automation** | Match worklist vs. DataGrid → generate CMS uploads + IndexOne inputs |
| 📋 | **Field Result Autostat** | Build daily autostat upload from the monthly field visit file |
| 📊 | **BAU Daily Status Report** | Merge MC3 DRR → SP MADRID formatted status report |

## How to Run Locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

Or double-click `run_app.bat` on Windows.

## Project Structure

```
ewb_portal/
├── app.py                  # Unified portal (entry point)
├── endo_automation.py      # Endorsement matching engine
├── requirements.txt
├── .gitignore
├── modules/
│   ├── endo_tab.py         # Endorsement Automation UI
│   ├── autostat_tab.py     # Field Result Autostat UI + logic
│   └── drr_tab.py          # BAU Daily Status Report UI + logic
└── ewb_data/               # ⚠️ Local only — NOT pushed to GitHub
    ├── status_map.csv           # Learned status code overrides
    ├── chcode_accounts.csv      # Learned CH code → account mappings
    ├── datagrid_accounts.csv    # DataGrid CH code lookups
    └── reports/                 # Dated copies of the BAU status report
```

## Deploy on Streamlit Community Cloud

1. Push this repo to GitHub (private recommended — contains internal tools).
2. Go to [share.streamlit.io](https://share.streamlit.io) → **New app**.
3. Select your repo, branch `main`, main file: `app.py`.
4. Click **Deploy**.

> ⚠️ **Data Privacy:** The `ewb_data/` folder and all `.xlsx` files are excluded by `.gitignore`.
> Never push borrower personal data to GitHub.

## Endorsement Automation Workflow

1. **Step 1**: Upload worklist + DataGrid → Run → download initial CMS CSV uploads.
2. **Step 1b**: Download the IndexOne input file → upload to IndexOne.
3. **Step 2**: Upload the IndexOne Mapped output → Ch Codes are filled → download final CMS uploads.

### Matching Rules
| Group | DataGrid Key |
|-------|-------------|
| BAU | ACCOUNT NUMBER |
| FFA / 150 DPD FFA / SPM REPO AI | CHASSIS (account no. only if chassis is blank/0) |

Priority: **BAU > FFA > 150 DPD > SPM REPO AI**

## Field Result Autostat Daily Routine

1. Drop the FIELD RESULT file (monthly workbook, sheet `RESULT`).
2. Pick the date (defaults to today).
3. The app filters `EWB AUTO LOAN RECOV`, removes CANCEL, deduplicates visits, resolves account numbers, and maps status codes.
4. Download the 6-column autostat file.

## BAU Daily Status Report Routine

1. *(Once)* Upload your last saved report in the sidebar.
2. Every day: drop today's MC3 DRR in the Upload tab — it is cleaned, merged, and saved automatically.
3. Press **Download status report**.
