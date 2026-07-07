// ============================================================
//  NBFC — Seller Document Collection Dashboard
//
//  READ-ONLY: this script never edits your data. It only reads
//  the existing master tab's headers and writes a single new tab
//  called "NBFC Dashboard" containing live formulas that mirror
//  the document-status columns already in the sheet.
//
//  SETUP:
//  1. Open the Google Sheet
//  2. Extensions -> Apps Script
//  3. Paste this file, Save (Ctrl+S)
//  4. Run  buildNbfcDashboard   (grant permissions when asked)
//     -> or use the "NBFC" menu that appears after reopening
// ============================================================

// Target spreadsheet. Leave '' to use the bound sheet.
const SHEET_ID = '1Ysw0VPLYcpcbIWngrLjeSYI-43sbxnNVlS3OJKhu2wE';

// The tab the dashboard writes to. Only this tab is created/rewritten.
const DASH_TAB = 'NBFC Dashboard';

// How the seller document block is located in the master tab.
// The block is the contiguous run of columns from the first to the
// last of these headers (matched case-insensitively, spaces collapsed).
const FIRST_DOC_HEADER = 'Bank Account Details';
const LAST_DOC_HEADER  = 'Last 6 Month platform sales ledger';

// Identity columns used for the seller-wise table.
const NAME_HEADER   = 'Seller Business Name';
const ENTITY_HEADER = 'Entity Type';
const REGION_HEADER = 'Region';

// Status vocabulary already used in the sheet.
//   Yes            -> received
//   NA / N/A       -> not applicable
//   No / - / blank -> pending
const RECEIVED_VALUES = ['Yes'];
const NA_VALUES       = ['NA', 'N/A'];

// Number of seller rows the dashboard mirrors (future-proof buffer).
const DASH_ROWS = 200;

// ── small helpers ────────────────────────────────────────────

function book_() {
  return SHEET_ID
    ? SpreadsheetApp.openById(SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

/** 1-based column index -> A1 letters. 1->A, 27->AA */
function colA1_(n) {
  let s = '';
  while (n > 0) {
    s = String.fromCharCode(64 + ((n - 1) % 26 + 1)) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function norm_(v) {
  return String(v == null ? '' : v).toLowerCase().replace(/\s+/g, ' ').trim();
}

/** Quote a sheet name for use in an A1 cross-sheet reference. */
function q_(name) {
  return "'" + String(name).replace(/'/g, "''") + "'";
}

/** COUNTIF over `range` for any of `values`, summed. */
function countAny_(range, values) {
  return values.map(v => `COUNTIF(${range},"${v}")`).join('+');
}

// ── locate the master tab + its columns (read only) ──────────

function locateSource_(ss) {
  const sheets = ss.getSheets();

  for (const sh of sheets) {
    if (sh.getName() === DASH_TAB) continue;
    const maxProbe = Math.min(sh.getLastRow(), 10);
    if (maxProbe < 1) continue;

    const width = Math.max(sh.getLastColumn(), 1);
    const probe = sh.getRange(1, 1, maxProbe, width).getValues();

    for (let r = 0; r < probe.length; r++) {
      const row = probe[r].map(norm_);
      const nameIdx = row.indexOf(norm_(NAME_HEADER));
      if (nameIdx === -1) continue;

      const headerRow = r + 1;                 // 1-based
      const find = h => row.indexOf(norm_(h)); // 0-based within row
      const firstDoc = find(FIRST_DOC_HEADER);
      const lastDoc  = find(LAST_DOC_HEADER);

      if (firstDoc === -1 || lastDoc === -1) {
        throw new Error(
          'Found the master tab "' + sh.getName() + '" but could not locate the ' +
          'document columns "' + FIRST_DOC_HEADER + '" .. "' + LAST_DOC_HEADER +
          '". Check those header names.'
        );
      }

      const docStart = Math.min(firstDoc, lastDoc) + 1; // 1-based
      const docEnd   = Math.max(firstDoc, lastDoc) + 1; // 1-based
      const labels   = probe[r].slice(docStart - 1, docEnd); // real header text

      return {
        sheet: sh,
        name: sh.getName(),
        headerRow: headerRow,
        dataStart: headerRow + 1,
        nameCol: nameIdx + 1,
        entityCol: (find(ENTITY_HEADER) + 1) || 0,
        regionCol: (find(REGION_HEADER) + 1) || 0,
        docStart: docStart,
        docEnd: docEnd,
        numDocs: docEnd - docStart + 1,
        docLabels: labels
      };
    }
  }

  throw new Error(
    'Could not find a tab containing the header "' + NAME_HEADER + '". ' +
    'Open the sheet and confirm the master data tab is present.'
  );
}

// ── main entry point ─────────────────────────────────────────

function buildNbfcDashboard() {
  const ss  = book_();
  const src = locateSource_(ss);

  const P    = q_(src.name) + '!';
  const dS   = colA1_(src.docStart);
  const dE   = colA1_(src.docEnd);
  const nCol = colA1_(src.nameCol);
  const eCol = src.entityCol ? colA1_(src.entityCol) : null;
  const N    = src.numDocs;

  const first = src.dataStart;
  const last  = src.dataStart + DASH_ROWS - 1;
  const block = `${P}${dS}${first}:${dE}${last}`;
  const nameRng = `${P}${nCol}${first}:${nCol}${last}`;

  // (Re)create the dashboard tab; never touch anything else.
  let sh = ss.getSheetByName(DASH_TAB);
  if (sh) { sh.clear(); sh.clearConditionalFormatRules(); }
  else    { sh = ss.insertSheet(DASH_TAB); }
  ss.setActiveSheet(sh);

  // Column widths
  const widths = [230, 150, 80, 80, 60, 70, 80, 130, 190];
  widths.forEach((w, i) => sh.setColumnWidth(i + 1, w));

  // ── Title ──
  sh.getRange('A1:I1').merge()
    .setValue('NBFC  —  Seller Document Collection Dashboard')
    .setBackground('#0D1B2A').setFontColor('#FFFFFF')
    .setFontSize(18).setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sh.setRowHeight(1, 46);

  sh.getRange('A2:I2').merge()
    .setFormula(`="Reflecting tab: ${src.name.replace(/"/g, '""')}"&"   |   Last updated: "&TEXT(NOW(),"dd-mmm-yyyy hh:mm")`)
    .setBackground('#1B2A3B').setFontColor('#90CAF9')
    .setFontSize(10).setHorizontalAlignment('center');
  sh.setRowHeight(2, 22);

  // ── Overall KPI strip (row 4 = labels, row 5 = values) ──
  const KPI_LBL = 4, KPI_VAL = 5;
  const sellers = `COUNTIF(${nameRng},"?*")`;
  const recv    = countAny_(block, RECEIVED_VALUES);
  const na      = countAny_(block, NA_VALUES);
  const appl    = `(${N}*${sellers}-(${na}))`;

  const kpis = [
    ['Total Sellers',   `=${sellers}`,                                             '#37474F', '#FFFFFF'],
    ['Docs Received',   `=${recv}`,                                                '#C8E6C9', '#1B5E20'],
    ['Docs Pending',    `=${appl}-(${recv})`,                                      '#FFCDD2', '#B71C1C'],
    ['Not Applicable',  `=${na}`,                                                  '#EEEEEE', '#616161'],
    ['Applicable',      `=${appl}`,                                                '#E3F2FD', '#0D47A1'],
    ['Completion',      `=IFERROR(TEXT((${recv})/(${appl}),"0.0%"),"—")`,          '#0D47A1', '#FFFFFF'],
  ];
  kpis.forEach((k, i) => {
    const c = i + 1;
    sh.getRange(KPI_LBL, c).setValue(k[0])
      .setBackground('#263238').setFontColor('#ECEFF1')
      .setFontSize(9).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
    sh.getRange(KPI_VAL, c).setFormula(k[1])
      .setBackground(k[2]).setFontColor(k[3])
      .setFontSize(15).setFontWeight('bold')
      .setHorizontalAlignment('center').setVerticalAlignment('middle');
  });
  sh.getRange(KPI_LBL, 7, 2, 3).setBackground('#FAFAFA'); // tidy empty G:I under strip
  sh.setRowHeight(KPI_LBL, 26);
  sh.setRowHeight(KPI_VAL, 40);

  // Legend
  sh.getRange('A6:I6').merge()
    .setValue('Legend:   Yes = received    ·    NA / N/A = not applicable    ·    No / "-" / blank = pending')
    .setBackground('#F5F5F5').setFontColor('#616161').setFontSize(9)
    .setHorizontalAlignment('center').setFontStyle('italic');
  sh.setRowHeight(6, 20);

  // ── Seller-wise table ──
  const SEC1 = 8, HDR1 = 9, DATA1 = 10;
  sh.getRange(SEC1, 1, 1, 9).merge()
    .setValue('▶  SELLER-WISE DOCUMENT STATUS')
    .setBackground('#1565C0').setFontColor('#FFFFFF')
    .setFontSize(12).setFontWeight('bold').setVerticalAlignment('middle');
  sh.setRowHeight(SEC1, 28);

  sh.getRange(HDR1, 1, 1, 9)
    .setValues([['Seller Business Name', 'Entity Type', 'Recv', 'Pend', 'N/A', 'Req.', '% Done', 'Status', 'Progress (of required)']])
    .setBackground('#37474F').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  sh.setRowHeight(HDR1, 24);

  const rows = [];
  for (let i = 0; i < DASH_ROWS; i++) {
    const r    = src.dataStart + i;
    const rng  = `${P}${dS}${r}:${dE}${r}`;
    const g    = `${P}${nCol}${r}=""`;
    const rc   = countAny_(rng, RECEIVED_VALUES);
    const nc   = countAny_(rng, NA_VALUES);
    const ap   = `(${N}-(${nc}))`;
    const pc   = `(${ap}-(${rc}))`;
    const ratio = `(${rc})/(${ap})`;

    rows.push([
      `=IF(${g},"",${P}${nCol}${r})`,
      eCol ? `=IF(${g},"",${P}${eCol}${r})` : '',
      `=IF(${g},"",${rc})`,
      `=IF(${g},"",${pc})`,
      `=IF(${g},"",${nc})`,
      `=IF(${g},"",${ap})`,
      `=IF(${g},"",IF(${ap}=0,"—",TEXT(${ratio},"0%")))`,
      `=IF(${g},"",IF(${ap}=0,"— no docs",IF(${pc}<=0,"✅ Complete","⏳ "&(${pc})&" pending")))`,
      `=IF(${g},"",IF(${ap}=0,"",REPT("█",ROUND(${ratio}*20,0))&REPT("░",20-ROUND(${ratio}*20,0))))`,
    ]);
  }
  sh.getRange(DATA1, 1, DASH_ROWS, 9).setFormulas(rows);
  sh.setRowHeights(DATA1, DASH_ROWS, 20);
  sh.getRange(DATA1, 3, DASH_ROWS, 5).setHorizontalAlignment('center');
  sh.getRange(DATA1, 8, DASH_ROWS, 1).setHorizontalAlignment('center');
  sh.getRange(DATA1, 9, DASH_ROWS, 1)
    .setFontFamily('Consolas').setFontSize(9).setFontColor('#1565C0');

  // ── Document-wise summary ──
  const SEC2 = DATA1 + DASH_ROWS + 1;
  const HDR2 = SEC2 + 1;
  const DATA2 = HDR2 + 1;

  sh.getRange(SEC2, 1, 1, 9).merge()
    .setValue('▶  DOCUMENT-WISE SUMMARY  (across all sellers)')
    .setBackground('#2E7D32').setFontColor('#FFFFFF')
    .setFontSize(12).setFontWeight('bold').setVerticalAlignment('middle');
  sh.setRowHeight(SEC2, 28);

  sh.getRange(HDR2, 1, 1, 6)
    .setValues([['Document', 'Received', 'Pending', 'N/A', '% Collected', 'Progress']])
    .setBackground('#37474F').setFontColor('#FFFFFF')
    .setFontWeight('bold').setFontSize(10).setHorizontalAlignment('center');
  sh.getRange(HDR2, 1).setHorizontalAlignment('left');
  sh.setRowHeight(HDR2, 24);

  const sellersCell = `$A$${KPI_VAL}`;   // Total Sellers
  const docRows = [];
  for (let i = 0; i < N; i++) {
    const c   = colA1_(src.docStart + i);
    const rng = `${P}${c}${first}:${c}${last}`;
    const rc  = countAny_(rng, RECEIVED_VALUES);
    const nc  = countAny_(rng, NA_VALUES);
    const ap  = `(${sellersCell}-(${nc}))`;
    const ratio = `(${rc})/(${ap})`;
    docRows.push([
      src.docLabels[i],
      `=${rc}`,
      `=${ap}-(${rc})`,
      `=${nc}`,
      `=IFERROR(TEXT(${ratio},"0%"),"—")`,
      `=IFERROR(REPT("█",ROUND(${ratio}*20,0))&REPT("░",20-ROUND(${ratio}*20,0)),"")`,
    ]);
  }
  sh.getRange(DATA2, 1, N, 1).setValues(docRows.map(r => [r[0]]));
  sh.getRange(DATA2, 2, N, 5).setFormulas(docRows.map(r => r.slice(1)));
  sh.setRowHeights(DATA2, N, 20);
  sh.getRange(DATA2, 2, N, 4).setHorizontalAlignment('center');
  sh.getRange(DATA2, 6, N, 1)
    .setFontFamily('Consolas').setFontSize(9).setFontColor('#2E7D32');

  // ── Conditional formatting ──
  const sellerStatus = sh.getRange(DATA1, 8, DASH_ROWS, 1);
  const sellerPct    = sh.getRange(DATA1, 7, DASH_ROWS, 1);
  const docPct       = sh.getRange(DATA2, 5, N, 1);
  sh.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('Complete').setBackground('#C8E6C9').setFontColor('#1B5E20')
      .setRanges([sellerStatus]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextContains('pending').setBackground('#FFF9C4').setFontColor('#E65100')
      .setRanges([sellerStatus]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('100%').setBackground('#C8E6C9').setFontColor('#1B5E20')
      .setRanges([sellerPct, docPct]).build(),
  ]);

  sh.setFrozenRows(HDR1);
  sh.setFrozenColumns(1);

  try {
    SpreadsheetApp.getUi().alert(
      '✅  Dashboard built.\n\n' +
      'Tab "' + DASH_TAB + '" now mirrors "' + src.name + '".\n' +
      'Found ' + N + ' document columns.\n\n' +
      'Nothing in your data was changed. Edit the master tab and\n' +
      'the dashboard recalculates automatically.'
    );
  } catch (e) { /* running without UI */ }
}

// Backwards-compatible alias.
function setupNBFCDashboard() { buildNbfcDashboard(); }

// Convenience menu (appears when the sheet is opened).
function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu('NBFC')
      .addItem('Rebuild Dashboard', 'buildNbfcDashboard')
      .addToUi();
  } catch (e) { /* no UI context */ }
}
