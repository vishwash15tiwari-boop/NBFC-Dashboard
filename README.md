# NBFC Document Dashboard

A Google Apps Script that builds a **read-only dashboard** on top of an existing
NBFC onboarding master sheet. It **does not change your data** — it only reads
the document-status columns already in the sheet and writes a single new tab,
**`NBFC Dashboard`**, filled with live formulas that recalculate automatically
whenever the master data changes.

**Target sheet:** `1Ysw0VPLYcpcbIWngrLjeSYI-43sbxnNVlS3OJKhu2wE`
(set in `SHEET_ID` at the top of `Code.gs`).

## What it reflects

The script locates your master tab (the one whose header row contains
`Seller Business Name`) and the contiguous block of seller document columns
between **`Bank Account Details`** and **`Last 6 Month platform sales ledger`**.
It reads the status values already in the sheet:

| Value in sheet | Meaning on dashboard |
|----------------|----------------------|
| `Yes` | Received |
| `NA` / `N/A` | Not applicable (excluded from the % denominator) |
| `No`, `-`, blank | Pending |

## What it builds (all on the `NBFC Dashboard` tab)

1. **Overall KPI strip** — Total Sellers, Docs Received, Docs Pending,
   Not Applicable, Applicable, Completion %.
2. **Seller-wise table** — one row per seller: received / pending / N/A counts,
   required (applicable) count, % done, a status flag
   (`✅ Complete` / `⏳ N pending`), and an ASCII progress bar.
3. **Document-wise summary** — for each document, how many sellers have provided
   it, how many are pending, how many N/A, and the % collected across all sellers
   (shows which documents are lagging).

Everything is formula-driven and references the master tab, so updates on the
sheet are reflected on the dashboard with no re-run needed.

## Setup — Option A: paste into the sheet (simplest)

1. Open the Google Sheet.
2. **Extensions → Apps Script**.
3. Replace the default code with `Code.gs` from this repo.
4. Save, then run `buildNbfcDashboard` and grant permissions.
5. Re-open the sheet to get an **NBFC → Rebuild Dashboard** menu.

## Setup — Option B: deploy with clasp (from this repo)

```bash
npm install -g @google/clasp
clasp login
clasp clone <SCRIPT_ID>              # or create a new standalone project
cp .clasp.json.example .clasp.json   # then fill in your scriptId
clasp push
```

`SHEET_ID` is already set to the target sheet, so a standalone deployment knows
which spreadsheet to read.

## Configuration (top of `Code.gs`)

| Constant | Purpose |
|----------|---------|
| `SHEET_ID` | Target spreadsheet id. Leave `''` to use the bound sheet. |
| `DASH_TAB` | Name of the tab the dashboard is written to (default `NBFC Dashboard`). |
| `FIRST_DOC_HEADER` / `LAST_DOC_HEADER` | Bounds of the document-column block in the master tab. |
| `NAME_HEADER` / `ENTITY_HEADER` / `REGION_HEADER` | Identity columns for the seller table. |
| `RECEIVED_VALUES` / `NA_VALUES` | Status vocabulary used in the sheet. |
| `DASH_ROWS` | How many seller rows to mirror (buffer for future rows). |

## Notes

- **Only** the `NBFC Dashboard` tab is created/rewritten. The master tab and all
  other tabs are left exactly as they are.
- Header matching is case-insensitive and whitespace-tolerant, so minor spacing
  differences in the master headers are handled.
- The buyer columns in this sheet are a buyer master list (name, rating, GST,
  PAN) with no document checklist, so the dashboard focuses on seller document
  collection. A buyer view can be added if a buyer document checklist is
  introduced.
