# PUF Bundle Plan Generator — Setup Guide

## Prerequisites

- Python 3.9+ installed
- Internet connection (for package install + Google Sheets)

---

## Step 1: Open Terminal / Command Prompt

Navigate to the project folder:

**Windows:**
```
cd C:\Users\Deepak\Desktop\resolved\puf_bundle_terminal - 24
```

**Mac:**
```
cd ~/Downloads/puf_bundle_terminal\ -\ 24
```

---

## Step 2: Install Dependencies

Run this command:

```
pip install -r requirements.txt
```

This installs: `pandas`, `openpyxl`, `streamlit`, `gspread`, `google-auth`

If `pip` doesn't work, try `pip3`:
```
pip3 install -r requirements.txt
```

---

## Step 3: Google Sheets Setup (One-Time)

### 3a. Create a Google Cloud Service Account

1. Go to https://console.cloud.google.com/
2. Create a new project (or use existing)
3. Go to **APIs & Services** → **Library** → Search **Google Sheets API** → **Enable**
4. Go to **APIs & Services** → **Credentials** → **Create Credentials** → **Service Account**
5. Give it a name (e.g. `puf-bundle-bot`) → Click **Create** → Skip optional steps → **Done**
6. Click on the service account you just created
7. Go to **Keys** tab → **Add Key** → **Create new key** → Select **JSON** → **Create**
8. A `.json` file will download — this is your credentials file

### 3b. Place the Credentials File

Rename the downloaded file to `service_account.json` and place it in the project root folder:

```
puf_bundle_terminal - 24/
├── bundle_plan_app.py
├── service_account.json    <-- place here
├── requirements.txt
└── ...
```

### 3c. Share Your Google Sheet

1. Open the `service_account.json` file in any text editor
2. Find the `"client_email"` field — it looks like: `puf-bundle-bot@your-project.iam.gserviceaccount.com`
3. Open your Google Sheet in browser
4. Click **Share** → Paste the service account email → Give **Editor** access → **Send**

---

## Step 4: Run the App

```
streamlit run bundle_plan_app.py
```

Or use the launcher:
```
python run_app.py
```

The app opens at http://localhost:8501

---

## Step 5: Configure Google Sheet URL in App

1. In the app sidebar, paste your Google Sheet URL in the **Sheet URL** field
2. The **Service Account JSON** path should auto-fill (if the file is in the project root)
3. If you placed it elsewhere, update the path in the sidebar

---

## Troubleshooting

| Issue | Fix |
|-------|-----|
| `gspread or google-auth not installed` | Run `pip install gspread google-auth` |
| `Service account file not found` | Place `service_account.json` in the project root folder |
| `Google Sheets failed: 403` | Share the sheet with the service account email (Step 3c) |
| `Google Sheets failed: SpreadsheetNotFound` | Check the Sheet URL is correct in the sidebar |
| Google Sheets not needed right now | Leave the Sheet URL empty — app works fully without it |

---

## Quick Reference (All Commands)

```
cd "path/to/puf_bundle_terminal - 24"
pip install -r requirements.txt
streamlit run bundle_plan_app.py
```

---

*Note: The app works fully without Google Sheets. Database saves locally regardless. Google Sheets is optional — configure it when ready.*
