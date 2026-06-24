# Sprinklr → Google Sheets Pipeline

A local automation tool that syncs Sprinklr social media exports directly into a reporting Google Sheet — replacing a manual copy-paste workflow with a config-driven, auditable pipeline.

Built with Python and Streamlit.

---

## What it does

Takes a Sprinklr export (`.xlsx`) and either:
- **Weekly Append** — adds new posts to the Google Sheet, skipping any permalink that already exists
- **Periodical Update** — refreshes metric columns for existing posts by matching on permalink

Both modes run a **dry-run preview first**, so you can review exactly what will change before writing anything to the Sheet.

---

## Architecture

```
Sprinklr Export (.xlsx)
        │
        ▼
Transformation Engine (process_sprinklr.py)
  - reads from row 3 (Sprinklr header offset)
  - maps Sprinklr column names → master column names
  - handles deleted posts, missing permalinks, duplicates
  - converts % fields, formats dates/times, sets NA for non-reel fields
        │
        ▼
Google Sheet (via gspread)
  - Weekly: appends new rows
  - Periodical: batch-updates metric cells only
        │
        ▼
Streamlit UI (app.py)
  - Upload → Dry Run → Review Audit → Confirm Write
```

---

## Folder structure

```
sprinklr-pipeline/
├── app.py                        # Streamlit app
├── process_sprinklr.py           # Core transformation engine
├── google_sheet_uploader.py      # gspread read/write logic
├── generate_token.py             # One-time Google OAuth setup
├── run_app.bat                   # Windows launcher
├── requirements.txt
├── client_secret_template.json   # Credential template (fill and rename)
├── config/
│   └── settings.json             # All configurable rules live here
├── input/
│   ├── weekly/                   # Drop weekly Sprinklr dumps here
│   └── periodical/               # Drop periodical Sprinklr dumps here
├── master/                       # Master workbook template
├── outputs/                      # Generated workbooks (gitignored)
└── logs/                         # JSON run logs (gitignored)
```

---

## Setup

### 1. Install dependencies

```bash
pip install -r requirements.txt
```

### 2. Set up Google Cloud credentials

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a project → enable the **Google Sheets API** and **Google Drive API**
3. Create an **OAuth 2.0 Client ID** (Desktop app type)
4. Download the credentials JSON
5. Copy `client_secret_template.json`, rename it to `client_secret.json`, and paste your credentials in

### 3. Authenticate

```bash
python generate_token.py
```

This opens a browser window for one-time OAuth sign-in. It creates `token.pickle` locally — never commit this file.

### 4. Configure

Edit `config/settings.json` to set:
- `sheet_url` — your Google Sheet URL
- `worksheet_name` — the tab name
- `sprinklr_to_master` — column name mappings
- `periodical_update_columns` — which columns to refresh in periodical mode

### 5. Run the app

```bash
run_app.bat
```

Or directly:

```bash
streamlit run app.py --server.port 8501
```

---

## Workflow

```
Upload Sprinklr dump
        │
        ▼
Choose mode: Weekly Append or Periodical Update
        │
        ▼
Click "Preview Changes" (dry run — nothing writes yet)
        │
        ▼
Review summary metrics + Row Audit table
        │
        ▼
Click "Update Google Sheet" to confirm write
```

---

## Row Audit

Every row in the Sprinklr dump is accounted for. Possible outcomes:

| Action | Reason |
|---|---|
| `appended` | New permalink, added to Sheet |
| `updated` | Existing permalink, metrics refreshed |
| `skipped` | Permalink already exists (weekly) or missing permalink |
| `removed` | Deleted post detected |
| `unmatched` | Permalink not found in Sheet (periodical only) |

---

## Configuration reference (`settings.json`)

| Key | Description |
|---|---|
| `sheet_url` | Full URL of the target Google Sheet |
| `worksheet_name` | Tab name within the Sheet |
| `sprinklr_header_row` | Row number where Sprinklr headers appear (default: 3) |
| `sprinklr_to_master` | Mapping of Sprinklr column names to master column names |
| `periodical_update_columns` | Columns to update in periodical mode |
| `percent_as_decimal_source_columns` | Sprinklr columns that store % as decimals |
| `numeric_source_columns` | Columns to coerce to numeric |
| `reel_permalink_contains` | String that identifies a reel permalink (default: `"reel"`) |
| `non_reel_override_columns` | Columns to set as `NA` for non-reel posts |
| `deleted_post_token` | String in Campaign Name that marks a deleted post |

---

## Security notes

- `client_secret.json` and `token.pickle` are gitignored — never commit them
- The pipeline reads from and writes to your Google Sheet using your own OAuth credentials
- No data is sent to any third-party service

---

## Tech stack

- Python 3.10+
- [Streamlit](https://streamlit.io/) — UI
- [gspread](https://github.com/burnash/gspread) — Google Sheets API client
- [openpyxl](https://openpyxl.readthedocs.io/) — Excel file processing
- [pandas](https://pandas.pydata.org/) — data handling
