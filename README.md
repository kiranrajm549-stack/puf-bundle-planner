# PUF Panel Bundle Plan Generator

Automates PUF (Polyurethane Foam) panel bundling for manufacturing.
Takes a Planning File and Job Card Export, cross-references them,
groups panels into physical bundles based on panel type/thickness rules,
and outputs a styled Excel report + printable QR-coded labels.

## Features
- Cross-references Planning File + Job Card Export by SO Number
- Applies 16 panel-type/thickness bundle size rules
- Longest panels bundled first (production-ready sequencing)
- Threshold-based split for oversized panels (large vs normal bundles)
- SQLite duplicate detection — flags already-planned orders across runs
- Session state caching — downloads don't re-run the pipeline
- Google Sheets auto-write via Service Account (optional)
- Outputs styled Excel bundle plan + printable QR-coded HTML labels

## Tech Stack
Python · Streamlit · pandas · openpyxl · SQLite · gspread

## How to Run
pip install -r requirements.txt
streamlit run bundle_plan_app.py

## Inputs Required
- Planning File (.xlsx) — SO Number, Material, Top Sheet, Region
- Job Card Export (.xlsx) — Job ID, Structure, Description, Length, Bal Qty

## Note
Do not upload service_account.json — create your own via Google Cloud Console.
See SETUP_GUIDE.md for full setup instructions.
