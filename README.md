# Recykal · Seller Onboarding — NBFC Document Tracker

A premium Google Apps Script web application (two files: **`Code.gs`** +
**`Index.html`**) that turns the *NBFC Document Tracker* workbook into a live,
management-friendly platform:

- **One-page executive dashboard** — portfolio completion, documents
  received/pending, document-wise collection status (sorted by most pending),
  sellers needing attention, region/entity breakdowns, and a full
  seller × document status matrix.
- **Seller entry & update** — add a new seller or update an existing one from a
  slide-over form. Saves write straight into the **`Seller_NBFC Tracker`** tab
  in its exact existing column order and vocabulary
  (`Received` / `Pending` / `NA` / free-text notes), and the dashboard refreshes
  in the same round trip.
- **Dynamic requirement checklist** — the mandatory-document matrix is read
  live from the **`Seller Requirement`** tab. Picking an entity type in the
  form automatically marks non-applicable documents `NA`
  (e.g. *Partnership Deed* for a Private Limited company). No document name is
  hardcoded anywhere — edit the sheet and the app follows.

## Data source

`CONFIG.SOURCE_FILE_ID` in `Code.gs` points at the Drive file
`1cCcBcFcoB0Dqj0_6uCQlIVOQomMd-Ayl` (*NBFC Document Tracker.xlsx*).

That file is an **uploaded Excel workbook**, which Apps Script cannot read or
write in place. On first run the script therefore converts it **once** into a
native Google Sheet — *"NBFC Document Tracker (Live)"*, created in the same
Drive folder with every tab and value intact — stores the new ID in Script
Properties, and uses it as the live backend from then on. The **Sheet** button
in the app's top bar always opens the live backend. If you later point
`SOURCE_FILE_ID` at a native Google Sheet, it is used directly and no copy is
made (clear the `BACKEND_SHEET_ID` script property after changing the ID).

## How statuses are interpreted

| Cell value in the tracker | Meaning |
|---|---|
| `Received` (or `Yes`, "…provided…") | Received |
| `NA` / `N/A` / `Not Applicable` | Not required for this entity type |
| `Pending`, blank, `-`, or any note containing "pending" | Pending (note shown in the UI) |

The `Pending Document` counter column is recomputed automatically on every
save, matching the sheet's existing convention (count of applicable documents
still pending).

## Deploy (5 minutes)

1. Open [script.google.com](https://script.google.com) → **New project**.
2. Replace the default `Code.gs` with this repo's `Code.gs`.
3. **+ File → HTML**, name it `Index`, paste this repo's `Index.html`.
4. (Optional) Project Settings → check *Show "appsscript.json"* and paste this
   repo's `appsscript.json` for the exact web-app config.
5. **Deploy → New deployment → Web app**:
   - *Execute as*: **Me**
   - *Who has access*: anyone you want using the tracker
6. Open the web-app URL. On the very first load, grant the requested
   permissions (Sheets + Drive — needed for the one-time xlsx conversion).

Alternatively push with [clasp](https://github.com/google/clasp):
`clasp create --type webapp && clasp push && clasp deploy`.

## Design

Recykal brand green (`#00A651`) with deep-ink neutrals, Inter type, and an
Apple/Notion-inspired layout: KPI strip with completion ring, thin rounded
bars, slide-over drawers, segmented controls, hover tooltips, skeleton loading,
and full mobile responsiveness. Status colors are consistent everywhere —
green = received, amber = pending, outlined gray = not applicable.
