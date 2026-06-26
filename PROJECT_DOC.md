# PUF Bundle Plan Generator

## Overview

Automates PUF (Polyurethane Foam) panel bundling for manufacturing. Takes a **Planning File** and **Job Card Export** (both `.xlsx`), cross-references them, groups panels into physical bundles based on panel type/thickness rules, and outputs a styled Excel report + printable QR-coded labels.

**Tech Stack**: Python, Streamlit (web UI), pandas, openpyxl, SQLite, gspread

---

## Flow Diagram

```
  PLANNING FILE (.xlsx)              JOB CARD EXPORT (.xlsx)
  ┌──────────────────┐              ┌──────────────────┐
  │ SO NUMBER        │              │ Job ID (=SO)     │
  │ MATERIAL         │              │ Structure        │
  │ TOP SHEET        │              │ Description      │
  │ REGION           │              │ Length / Width    │
  │ CUSTOMER         │              │ Bal Qty          │
  └────────┬─────────┘              └────────┬─────────┘
           │                                  │
           └────────────┬─────────────────────┘
                        ▼
              PARSE & CROSS-REFERENCE
              • Join on SO Number
              • Match panel type + thickness
              • Extract colour from TOP SHEET
              • Detect exceptions
                        │
                        ▼
                  DEDUP CHECK (SQLite)
              • Compare (WORK_ID, COLOUR)
              • Skip already-planned orders
              • Flag duplicates in UI
                        │
                        ▼
                BUNDLE GENERATION
              • Longest panels bundled first
              • Layer rules per type/thickness
              • Threshold split for oversized
              • Sequential Bundle IDs
                        │
              ┌─────────┼─────────┐
              ▼                   ▼
        EXCEL REPORT        HTML LABELS
     • BUNDLE PLAN sheet    • 4×6" per bundle
     • ORDER DATA sheet     • QR code
     • PENDING sheet        • Panel matrix
              │                   │
              └─────────┬─────────┘
                        ▼
              SESSION STATE CACHE
           (downloads don't rerun pipeline)
                        │
                        ▼
              CONFIRM & SAVE
              • SQLite DB write
              • Google Sheets append
```

---

## Key Files

| File | Purpose |
|------|---------|
| `bundle_plan_app.py` | Main Streamlit app (all logic) |
| `run_app.py` | Headless launcher |
| `bundle_plan.db` | SQLite dedup database (auto-created) |
| `service_account.json` | Google Sheets credentials (user-provided) |
| `requirements.txt` | Dependencies |
| `PUFBundleApp.spec` | PyInstaller packaging spec |

---

## Bundle Rules (LAYER_DATA)

| Panel Type | Bundle Size | Threshold | Large Size |
|-----------|------------|-----------|------------|
| 30MM WALL | 36 | 10000mm | 10 |
| 50MM WALL | 22 | 10000mm | 10 |
| 30MM ROOF | 22 | 8000mm | 12 |
| 50MM ROOF | 16 | 8000mm | 10 |
| Others | 7–26 | None | None |

Panels longer than threshold get a smaller bundle size. All panels are bundled **longest first**.

---

## Resolutions Applied (v2)

### 1. Bundle Length Ordering — FIXED
**Problem**: Bundles were sequenced shortest-length-first. In production, longest panels are always cut first.
**Fix**: Large-length (above threshold) bundles are now created first, getting lower sequence numbers (Bundle-01, 02...). Normal-length panels follow.

### 2. Duplicate Planning Detection — ADDED
**Problem**: No way to detect if an order was already planned in a previous run.
**Fix**: SQLite database (`bundle_plan.db`) stores every confirmed `(WORK_ID, COLOUR)` pair. On "Generate", existing entries are checked and duplicates are flagged/skipped.

### 3. Download Restart Bug — FIXED
**Problem**: Clicking any download button reran the entire pipeline. With dedup active, this caused the just-generated data to be flagged as duplicate — making downloads impossible.
**Fix**: Results are cached in `st.session_state`. Download clicks re-display cached data without re-running the pipeline. DB write happens only on explicit "Confirm & Save".

### 4. Google Sheets Auto-Write — ADDED
**Problem**: Manual data entry into tracking sheet after each run.
**Fix**: On "Confirm & Save", bundle rows are auto-appended to a configured Google Sheet via Service Account auth. Falls back gracefully if Sheets API fails — local DB always saves.

---

## User Workflow

```
1. Upload Planning File + Job Card Export
2. Click [Generate Bundle Plan]
3. Review results → download Excel / Labels freely
4. Click [Confirm & Save] → writes to DB + Google Sheets
5. Future runs will detect already-planned orders
```

---

## Google Sheets Setup (One-Time)

1. Google Cloud Console → Create Service Account → Enable Sheets API
2. Download JSON key → save as `service_account.json` in project root
3. Share your Google Sheet with the service account email
4. Paste Sheet URL in sidebar

---

*Last updated: 27 April 2026*
