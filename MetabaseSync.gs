/* ═══════════════════════════════════════════════════════════════════════════
   Metabase → Google Sheets Sync  (Near Real-Time)
   meta.recykal.com  ·  Query 5712 → Sellers tab  |  Query 5711 → Buyers tab
   ─────────────────────────────────────────────────────────────────────────
   SETUP (do once):
     1. Find METABASE_PASS below → replace YOUR_PASSWORD_HERE with your password
     2. Save (Ctrl + S)
     3. Run createAutoSync() once → sheet auto-refreshes every minute

   TO STOP:   run deleteAutoSync()
   MANUAL RUN: run syncMetabaseToSheet()
   DEBUG:      run debugColumns() to see exact column names Metabase returns
   ═══════════════════════════════════════════════════════════════════════════ */

// ─────────────────────────────────────────────────────────────────────────────
// CONFIGURATION
// ─────────────────────────────────────────────────────────────────────────────

var CFG = {
  METABASE_URL    : 'https://meta.recykal.com',
  METABASE_USER   : 'vishwash.tiwari@recykal.com',
  METABASE_PASS   : 'YOUR_PASSWORD_HERE',   // ← replace with your password
  SHEET_ID        : '1k5-2k__JnwEgGWJ8s5xRuKbnTvqam76k-TeBIDiN7os',
  SYNC_EVERY_MINS : 1,

  /* One entry per Metabase card. `gstinKey` / `nameKey` are FIELD_MAP keys used
     to locate those two columns in the query output, so a renamed Metabase
     column is handled by extending FIELD_MAP rather than editing logic here. */
  QUERIES: [
    { id: 5712, tab: 'Sellers', gstinKey: 'sellergstin', nameKey: 'sellerbusinessname' },
    { id: 5711, tab: 'Buyers',  gstinKey: 'buyergstin',  nameKey: 'buyerbusinessname'  },
  ],

  /* Only Open Marketplace records are synced. ONBOARDING_STATUS is disabled
     (empty string) so vertical is the sole filter; set it back to 'Completed'
     to also require a finished onboarding. */
  FILTER: {
    VERTICAL         : 'Open Marketplace',
    ONBOARDING_STATUS: '',
  },

  /* Header aliases used to locate the three columns this sync touches in the
     destination tab. Nothing is written outside these columns. */
  SHEET_COLS: {
    no:    ['No.', 'No', 'S. No.', 'S No', 'Sr No', 'Sr. No.', 'Serial No', 'Serial',
            'SNo', '#', 'Sl No', 'Sl. No.'],
    gstin: ['GSTIN', 'Seller GSTIN', 'Buyer GSTIN', 'GSTIN No', 'GSTIN Number',
            'GST Number', 'GST No', 'Seller GST Number', 'Buyer GST Number',
            'Seller GSTIN Number', 'Buyer GSTIN Number', 'GST'],
    name:  ['Seller Business Name', 'Buyer Business Name', 'Seller Name', 'Buyer Name',
            'Entity Name', 'Business Name', 'Company Name', 'Legal Name', 'Trade Name',
            'Vendor Name', 'Party Name', 'Name'],
  },

  /* Last-resort column positions (1-based) used only when a header cannot be
     resolved by name. Matches the confirmed layout: A = No., B = GSTIN,
     C = Seller/Buyer Name. A fallback is refused if the header already at that
     position clearly belongs to something else, so data is never overwritten. */
  FALLBACK_COLS: { no: 1, gstin: 2, name: 3 },

  /* Row 1 is normally the header. If a tab has a title/banner row instead, the
     first this many rows are scanned for the row that holds a GSTIN header. */
  HEADER_SCAN_ROWS: 8,
};

// ─────────────────────────────────────────────────────────────────────────────
// SCHEMA
// ─────────────────────────────────────────────────────────────────────────────

// Exact column order that must appear in the sheet (header row 1).
// Used only when the tab has no existing headers — otherwise row 1 of the
// sheet is the source of truth.  Column names here match the Sellers tab;
// the Buyers tab pre-exists with 'Buyer Business Name' / 'Buyer Type'.
var COLUMN_ORDER = [
  'No.',
  'Seller Business Name',
  'Region',
  'Vertical',
  'Seller Type',
  'State',
  'GSTIN',
  'Vintage with Recykal',
  'Entity Type',
];

// Document columns — values are mapped: 1/2/3 → "Received", 0 → "Not Received",
// column absent from query output → "NA"
var DOC_COLUMNS = [
  '2 yr Audited Financial, Current Provisional',
  'Bank Statement',
  'GSTR 3B - 12 Month',
  'ITR (Last 2 years)',
  'CIBIL Consent',
  'Sanction Letter of all loans',
  'Partnership Deed',
  'Entity PAN',
  'MSME',
  'GST Certificate',
  'Owner / Director / Partner PAN',
  'Aadhar (Owner / Director / Partner)',
  'Electricity Bill / Rental Agreement',
  'MOA, AOA, COI',
  'Shareholding Details',
];
// Normalised set for O(1) lookup
var DOC_COL_NORMS = {};
DOC_COLUMNS.forEach(function (c) { DOC_COL_NORMS[norm_(c)] = true; });

// Explicit sheet-header → Metabase column name aliases.
// Keys are the normalised sheet header (via norm_()).
// Values are arrays of possible Metabase display_name strings, tried in order.
// This covers cases where the two sides use completely different wording.
// Maps normalised sheet header → ordered list of Metabase column name candidates.
// norm_() collapses spaces, underscores, hyphens, etc., so "Entity Name", "entity_name",
// and "EntityName" all normalise to "entityname" — one entry covers all three variants.
var FIELD_MAP = {
  // ── Business name — first candidate wins; others are safe fallbacks ──────────
  // Each list ends with the *other* side's wording. The Buyers tab is seeded
  // with the seller header text ("Seller Business Name"), so a buyer query that
  // returns "Buyer Name" must still resolve — that mismatch is exactly why the
  // buyer name column came through empty.
  'sellerbusinessname' : [
    'Entity Name',          // confirmed by user
    'Seller Name', 'Seller Business Name',
    'Name', 'Business Name', 'Company Name', 'Organisation Name', 'Organization Name',
    'Vendor Name', 'Party Name', 'Legal Name', 'Trade Name',
    'Buyer Name', 'Buyer Business Name', 'Customer Name',
  ],
  'buyerbusinessname'  : [
    'Entity Name',
    'Buyer Name',  'Buyer Business Name', 'Customer Name', 'Customer Business Name',
    'Name', 'Business Name', 'Company Name', 'Organisation Name', 'Organization Name',
    'Vendor Name', 'Party Name', 'Legal Name', 'Trade Name',
    'Seller Name', 'Seller Business Name',
  ],

  // ── Region — use Meta's Region column if present, else derive from State ─────
  'region' : ['Region', 'region', 'State'],

  // ── Core identity ─────────────────────────────────────────────────────────────
  'vertical'    : ['Vertical', 'Business Vertical', 'Biz Vertical'],

  // Legal structure — "Partnership", "Proprietorship", "Private Limited", etc.
  'entitytype'  : [
    'Entity Type', 'Entity_Type', 'Legal Entity Type',
    'Legal Structure', 'Organisation Type', 'Organization Type',
    'Type of Entity', 'Registration Type',
  ],

  // Business category — "Baler", "Trader", "Manufacturer", "Baler cum Trader", etc.
  'sellertype'  : [
    'Seller Type', 'Seller_Type',
    'Seller Category', 'Category', 'Business Type', 'Business Category', 'Type',
  ],
  'buyertype'   : [
    'Buyer Type',  'Buyer_Type',
    'Buyer Category',  'Category', 'Business Type', 'Business Category', 'Type',
  ],

  // ── Contact / ID fields ───────────────────────────────────────────────────────
  'state'       : ['State'],
  'sellergstin' : ['GSTIN', 'Seller GSTIN', 'GST Number', 'GST No', 'GST', 'Seller_GSTIN'],
  'buyergstin'  : ['GSTIN', 'Buyer GSTIN',  'GST Number', 'GST No', 'GST', 'Buyer_GSTIN'],
  'vintagewithrecykal' : ['Vintage with Recykal', 'Vintage_with_Recykal', 'Vintage'],
  'finoscalerating'    : [
    'Finoscale Rating', 'Finoscale_Rating', 'Finoscale Score',
    'Credit Rating', 'Rating',
  ],
  'mailid'   : ['Mail ID', 'Mail_ID', 'Email ID', 'Email', 'Mail', 'Email Address'],
  'pocmail'  : ['POC Mail', 'POC_mail', 'POC Email', 'POC Mail ID', 'POC Email ID'],
  'mobileno' : ['Mobile No', 'Mobile_No', 'Mobile Number', 'Mobile', 'Phone', 'Contact No'],
  'pocnames' : ['POC Names', 'POC Name', 'Point of Contact', 'POC'],
  'dateofregistration' : [
    'Effective Date of Registration', 'Effective_Date_Of_Registration',
    'Date of Registration', 'Date_of_Registration',
    'Registration Date', 'Reg Date', 'Incorporation Date',
    'Onboarding Date', 'Date of Onboarding', 'Joining Date', 'Joined Date',
    'Platform Registration Date', 'Seller Registration Date', 'Buyer Registration Date',
  ],
  'pendingdocument' : ['Pending Document', 'Pending_Document', 'Pending Documents'],
  'debtprofile'     : ['Debt Profile', 'Debt_Profile', 'Debt'],
};


// ─────────────────────────────────────────────────────────────────────────────
// MAIN ENTRY POINT
// ─────────────────────────────────────────────────────────────────────────────

/**
 * Sync Open Marketplace Sellers and Buyers from Metabase into the sheet.
 *
 * Append-only and idempotent. For each tab the sync reads the existing GSTINs,
 * appends only the records that are not already present, and renumbers the
 * "No." column. Existing rows are never rewritten, and the only columns ever
 * written are No. / GSTIN / Name — every other column (documents, vintage,
 * eligibility, remarks, formulas, formatting) is left untouched.
 *
 * A script lock serialises runs so an overlapping trigger cannot append the
 * same GSTIN twice.
 */
function syncMetabaseToSheet() {
  if (!CFG.METABASE_PASS || CFG.METABASE_PASS === 'YOUR_PASSWORD_HERE') {
    throw new Error('Set CFG.METABASE_PASS to your Metabase password and save the script.');
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    Logger.log('⏭ Another sync is already running — skipped (no duplicates created).');
    return;
  }

  var t0 = Date.now();
  try {
    var token = getSessionToken_();
    var ss    = SpreadsheetApp.openById(CFG.SHEET_ID);
    var total = { added: 0, existing: 0, blank: 0 };

    CFG.QUERIES.forEach(function (q) {
      Logger.log('━━━ Query ' + q.id + ' → "' + q.tab + '" ━━━');
      try {
        var raw = fetchCardData_(token, q.id);
        Logger.log('Fetched: ' + raw.rows.length + ' rows, ' + raw.cols.length + ' cols');
        Logger.log('Metabase columns: ' + raw.cols.join(' | '));

        var filtered = applyFilter_(raw);
        Logger.log('After Open Marketplace filter: ' + filtered.rows.length + ' rows');
        if (filtered.rows.length === 0) {
          Logger.log('⚠ Zero rows after filter — run debugColumns() and check CFG.FILTER');
        }

        var r = appendNewEntities_(ss, q, filtered);
        total.added    += r.added;
        total.existing += r.existing;
        total.blank    += r.blank;
      } catch (e) {
        // One failing card must not stop the other from syncing.
        Logger.log('✗ "' + q.tab + '" failed: ' + (e && e.message || e));
      }
    });

    Logger.log('══ Sync complete in ' + (Math.round((Date.now() - t0) / 100) / 10) + 's — ' +
               total.added + ' added, ' + total.existing + ' already present, ' +
               total.blank + ' skipped (no GSTIN) ══');
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/* ─────────────────────────────────────────────────────────────────────────────
   APPEND-ONLY WRITER
   ───────────────────────────────────────────────────────────────────────────*/

/** Index of the first column in `cols` matching any candidate name, else -1. */
function mbFindIdx_(cols, candidates) {
  // Alias priority, not column order: an exact "GSTIN" wins over a stray "GST".
  for (var j = 0; j < candidates.length; j++) {
    var want = norm_(candidates[j]);
    if (!want) continue;
    for (var i = 0; i < cols.length; i++) if (norm_(cols[i]) === want) return i;
  }
  return -1;
}

/** A1-style letter for a 0-based column index, for readable logs. */
function mbColLetter_(idx) {
  var s = '', n = idx + 1;
  while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

/* Fuzzy matchers. Each returns a column index or -1, and each carries explicit
   exclusions so a document column can never be mistaken for an identity column
   ("GST Certificate"/"GSTR 3B" are not GSTIN; "POC Name" is not the entity name). */
var MB_FUZZY = {
  gstin: function (n) {
    if (n.indexOf('gst') === -1) return false;
    if (/certificate|return|gstr|3b|filing|status|doc/.test(n)) return false;
    return n === 'gst' || n === 'gstin' || /gstin|gstno|gstnumber/.test(n);
  },
  name: function (n) {
    if (n.indexOf('name') === -1) return false;
    if (/poc|contact|person|bank|director|owner|partner|promoter|user|login|father|authorised|authorized/.test(n)) return false;
    return true;
  },
  no: function (n) {
    return n === 'no' || n === 'no.' || n === 'sno' || n === 'srno' || n === 'slno' ||
           n === 'serial' || n === 'serialno' || n === '#';
  }
};

/* Headers that must never be overwritten by a positional fallback. */
function mbHeaderLooksReserved_(header) {
  var n = norm_(header);
  if (!n) return false;                       // blank header is safe to claim
  return /gst|document|doc|bank|vintage|eligib|remark|status|state|region|vertical|type|date|pan|aadhar|aadhaar|itr|cibil|sanction|msme|moa|aoa|coi|shareholding|ledger|kyc|constitution|audit|debt|electricity|rental|deed|email|mobile|phone/.test(n);
}

/**
 * Resolve one logical column in a sheet header row.
 * Order: exact alias → guarded fuzzy → configured position (only if the header
 * sitting there is blank or not obviously something else).
 * @return {{idx:number, how:string}} idx is 0-based, or -1 when unresolved.
 */
function mbResolveSheetCol_(headers, key, tabName) {
  var idx = mbFindIdx_(headers, CFG.SHEET_COLS[key] || []);
  if (idx >= 0) return { idx: idx, how: 'header "' + headers[idx] + '"' };

  var fuzzy = MB_FUZZY[key];
  if (fuzzy) {
    for (var i = 0; i < headers.length; i++) {
      if (fuzzy(norm_(headers[i]))) return { idx: i, how: 'fuzzy header "' + headers[i] + '"' };
    }
  }

  var pos = (CFG.FALLBACK_COLS || {})[key];
  if (pos) {
    var p = pos - 1;
    var at = headers.length > p ? headers[p] : '';
    if (!mbHeaderLooksReserved_(at)) {
      return { idx: p, how: 'position ' + mbColLetter_(p) + (at ? ' (header "' + at + '")' : ' (blank header)') };
    }
    Logger.log('  ⚠ "' + tabName + '": refused ' + key + ' fallback to column ' +
               mbColLetter_(p) + ' — it holds "' + at + '"');
  }
  return { idx: -1, how: 'unresolved' };
}

/**
 * Locate the header row. Normally 1; scans further only when row 1 has no
 * GSTIN-like header (a tab with a title/banner row above the real headers).
 * @return {number} 1-based row number.
 */
function mbFindHeaderRow_(sh) {
  var scan = Math.min(CFG.HEADER_SCAN_ROWS || 1, Math.max(1, sh.getLastRow()));
  var cols = Math.max(1, sh.getLastColumn());
  if (scan < 1) return 1;
  var block = sh.getRange(1, 1, scan, cols).getValues();
  for (var r = 0; r < block.length; r++) {
    if (mbFindIdx_(block[r], CFG.SHEET_COLS.gstin) >= 0) return r + 1;
    for (var c = 0; c < block[r].length; c++) if (MB_FUZZY.gstin(norm_(block[r][c]))) return r + 1;
  }
  return 1;
}

/** GSTIN comparison key — trimmed and upper-cased, matching the rest of the app. */
function mbGstinKey_(v) { return String(v == null ? '' : v).trim().toUpperCase(); }

/**
 * Append the Open Marketplace records whose GSTIN is not already in the tab,
 * then renumber "No.". Returns { added, existing, blank }.
 */
function appendNewEntities_(ss, q, result) {
  var out = { added: 0, existing: 0, blank: 0 };

  // ── Locate GSTIN + Name in the Metabase output (alias, then guarded fuzzy) ──
  var gIdx = mbFindIdx_(result.cols, FIELD_MAP[q.gstinKey] || []);
  if (gIdx < 0) for (var gi = 0; gi < result.cols.length; gi++) {
    if (MB_FUZZY.gstin(norm_(result.cols[gi]))) { gIdx = gi; break; }
  }
  var nIdx = mbFindIdx_(result.cols, FIELD_MAP[q.nameKey] || []);
  if (nIdx < 0) for (var ni = 0; ni < result.cols.length; ni++) {
    if (MB_FUZZY.name(norm_(result.cols[ni]))) { nIdx = ni; break; }
  }

  if (gIdx < 0) {
    Logger.log('✗ No GSTIN column in query output — nothing appended. Columns: ' + result.cols.join(' | '));
    return out;
  }
  Logger.log('  Metabase → GSTIN: "' + result.cols[gIdx] + '"' +
             (nIdx >= 0 ? ', Name: "' + result.cols[nIdx] + '"' : ''));
  if (nIdx < 0) {
    Logger.log('  ⚠ No name column found in query output — names will be blank. Columns: ' +
               result.cols.join(' | ') + '  → add the correct header to FIELD_MAP.' + q.nameKey);
  }

  // ── Target tab, created with headers only if it does not exist ───────────
  var sh = ss.getSheetByName(q.tab);
  if (!sh) {
    sh = ss.insertSheet(q.tab);
    sh.getRange(1, 1, 1, 3).setValues([['No.', 'GSTIN',
      q.tab === 'Buyers' ? 'Buyer Business Name' : 'Seller Business Name']]);
    sh.setFrozenRows(1);
    Logger.log('  Created tab "' + q.tab + '" with No. / GSTIN / Name headers.');
  }

  var headerRow = mbFindHeaderRow_(sh);
  var lastRow   = sh.getLastRow();
  var lastCol   = Math.max(1, sh.getLastColumn());
  var headers   = sh.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  Logger.log('  Header row ' + headerRow + ': ' + headers.join(' | '));

  var rNo    = mbResolveSheetCol_(headers, 'no',    q.tab);
  var rGstin = mbResolveSheetCol_(headers, 'gstin', q.tab);
  var rName  = mbResolveSheetCol_(headers, 'name',  q.tab);
  var cNo = rNo.idx, cGstin = rGstin.idx, cName = rName.idx;

  if (cGstin < 0) {
    Logger.log('✗ "' + q.tab + '": GSTIN column unresolved — nothing appended. Headers: ' + headers.join(' | '));
    return out;
  }
  Logger.log('  Sheet → GSTIN: ' + mbColLetter_(cGstin) + ' via ' + rGstin.how);
  Logger.log('  Sheet → Name : ' + (cName >= 0 ? mbColLetter_(cName) + ' via ' + rName.how : 'UNRESOLVED — names will not be written'));
  Logger.log('  Sheet → No.  : ' + (cNo   >= 0 ? mbColLetter_(cNo)   + ' via ' + rNo.how   : 'UNRESOLVED — numbering skipped'));

  // ── Existing GSTINs: one read of the key column only ─────────────────────
  var firstDataRow = headerRow + 1;
  var existing = {};
  var dataRows = Math.max(0, lastRow - headerRow);
  var gstinCol = dataRows > 0 ? sh.getRange(firstDataRow, cGstin + 1, dataRows, 1).getValues() : [];
  for (var r = 0; r < gstinCol.length; r++) {
    var k = mbGstinKey_(gstinCol[r][0]);
    if (k) existing[k] = true;
  }
  var uniqueExisting = Object.keys(existing).length;
  Logger.log('  Existing rows: ' + dataRows + ' (' + uniqueExisting + ' unique GSTINs)');
  if (dataRows > uniqueExisting) {
    Logger.log('  ⚠ ' + (dataRows - uniqueExisting) + ' duplicate/blank row(s) already in "' + q.tab +
               '" — run removeDuplicateGstins("' + q.tab + '") to clean up.');
  }

  // ── Collect the genuinely new records (deduped within the batch too) ─────
  var seen = {}, newRows = [];
  result.rows.forEach(function (row) {
    var key = mbGstinKey_(row[gIdx]);
    if (!key)          { out.blank++;    return; }
    if (existing[key]) { out.existing++; return; }   // already in the sheet → skip
    if (seen[key])     { out.existing++; return; }   // duplicate inside this fetch
    seen[key] = true;
    newRows.push({
      gstin: String(row[gIdx]).trim(),
      name:  nIdx >= 0 ? String(row[nIdx] == null ? '' : row[nIdx]).trim() : ''
    });
  });

  // ── Append, writing ONLY the GSTIN and Name columns ──────────────────────
  if (newRows.length) {
    var startRow = lastRow + 1;
    var needed   = startRow + newRows.length - 1;
    if (needed > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), needed - sh.getMaxRows());

    sh.getRange(startRow, cGstin + 1, newRows.length, 1)
      .setValues(newRows.map(function (x) { return [x.gstin]; }));
    if (cName >= 0) {
      sh.getRange(startRow, cName + 1, newRows.length, 1)
        .setValues(newRows.map(function (x) { return [x.name]; }));
    }
    out.added = newRows.length;
    Logger.log('  ✓ Appended ' + newRows.length + ' new record(s) at row ' + startRow);
  } else {
    Logger.log('  ✓ Nothing new — sheet already current');
  }

  // Numbering runs even when nothing was appended, so gaps get repaired.
  if (cNo >= 0) renumberNoColumn_(sh, cNo, cGstin, headerRow);
  Logger.log('  Summary: +' + out.added + ' added, ' + out.existing +
             ' already present, ' + out.blank + ' without GSTIN');
  return out;
}

/**
 * Renumber the "No." column 1..N over every row that has a GSTIN.
 * Written only when the current numbering is already wrong, so a no-op sync
 * performs no write at all.
 */
function renumberNoColumn_(sh, cNo, cGstin, headerRow) {
  headerRow = headerRow || 1;
  var firstDataRow = headerRow + 1;
  var dataRows = sh.getLastRow() - headerRow;
  if (dataRows < 1) return;

  var gstins  = sh.getRange(firstDataRow, cGstin + 1, dataRows, 1).getValues();
  var current = sh.getRange(firstDataRow, cNo    + 1, dataRows, 1).getValues();

  var next = 1, desired = [], changed = false;
  for (var i = 0; i < dataRows; i++) {
    var want = mbGstinKey_(gstins[i][0]) ? next++ : '';     // blank rows stay unnumbered
    desired.push([want]);
    if (String(current[i][0] == null ? '' : current[i][0]) !== String(want)) changed = true;
  }
  if (!changed) { Logger.log('  Numbering already correct — no write'); return; }

  sh.getRange(firstDataRow, cNo + 1, dataRows, 1).setValues(desired);
  Logger.log('  ✓ Renumbered "No." 1–' + (next - 1));
}

/* ─────────────────────────────────────────────────────────────────────────────
   DIAGNOSTICS + DUPLICATE REPAIR
   ───────────────────────────────────────────────────────────────────────────*/

/**
 * Run this first when anything looks wrong. Prints, for both tabs, the detected
 * header row with column letters and exactly which column each lookup resolved
 * to — plus the columns each Metabase card returns. Reads only; writes nothing.
 */
function debugSheetColumns() {
  var ss = SpreadsheetApp.openById(CFG.SHEET_ID);
  Logger.log('════ SHEET: ' + ss.getName() + ' ════');

  CFG.QUERIES.forEach(function (q) {
    var sh = ss.getSheetByName(q.tab);
    Logger.log('\n──── Tab "' + q.tab + '" ────');
    if (!sh) { Logger.log('  ✗ tab not found'); return; }

    var headerRow = mbFindHeaderRow_(sh);
    var lastCol   = Math.max(1, sh.getLastColumn());
    var headers   = sh.getRange(headerRow, 1, 1, lastCol).getValues()[0];
    Logger.log('  Rows: ' + sh.getLastRow() + '  Cols: ' + lastCol + '  Header row: ' + headerRow);
    headers.forEach(function (h, i) {
      if (String(h).trim() !== '') Logger.log('    ' + mbColLetter_(i) + ': "' + h + '"');
    });

    ['no', 'gstin', 'name'].forEach(function (k) {
      var r = mbResolveSheetCol_(headers, k, q.tab);
      Logger.log('  → ' + k + ': ' + (r.idx >= 0 ? mbColLetter_(r.idx) + ' via ' + r.how : 'UNRESOLVED'));
    });

    var rG = mbResolveSheetCol_(headers, 'gstin', q.tab);
    if (rG.idx >= 0 && sh.getLastRow() > headerRow) {
      var vals = sh.getRange(headerRow + 1, rG.idx + 1, sh.getLastRow() - headerRow, 1).getValues();
      var seen = {}, dups = 0, blanks = 0;
      vals.forEach(function (row) {
        var k = mbGstinKey_(row[0]);
        if (!k) { blanks++; return; }
        if (seen[k]) dups++; else seen[k] = true;
      });
      Logger.log('  GSTINs: ' + Object.keys(seen).length + ' unique, ' + dups + ' duplicate row(s), ' + blanks + ' blank');
      if (dups) Logger.log('  → run removeDuplicateGstins("' + q.tab + '") to clean up');
    }
  });

  var token = getSessionToken_();
  CFG.QUERIES.forEach(function (q) {
    Logger.log('\n──── Metabase card ' + q.id + ' (' + q.tab + ') ────');
    try {
      var raw = fetchCardData_(token, q.id);
      Logger.log('  ' + raw.rows.length + ' rows, columns:');
      raw.cols.forEach(function (c) { Logger.log('    "' + c + '"'); });
      var g = mbFindIdx_(raw.cols, FIELD_MAP[q.gstinKey] || []);
      var n = mbFindIdx_(raw.cols, FIELD_MAP[q.nameKey]  || []);
      Logger.log('  → GSTIN: ' + (g >= 0 ? '"' + raw.cols[g] + '"' : 'UNRESOLVED'));
      Logger.log('  → Name : ' + (n >= 0 ? '"' + raw.cols[n] + '"' : 'UNRESOLVED'));
      Logger.log('  After Open Marketplace filter: ' + applyFilter_(raw).rows.length + ' rows');
    } catch (e) { Logger.log('  ✗ ' + (e && e.message || e)); }
  });
  Logger.log('\n════ end ════');
}

/** Report duplicate GSTIN rows without changing anything. */
function previewDuplicateGstins(tabName) { return mbDedupe_(tabName, true); }

/**
 * Remove duplicate GSTIN rows, keeping the FIRST occurrence of each — the
 * original row that carries the document, vintage and eligibility data. Later
 * copies (the ones a bad sync appended) are deleted bottom-up so row indices
 * stay valid. Call previewDuplicateGstins() first to see what would go.
 * @param {string=} tabName  omit to process every configured tab.
 */
function removeDuplicateGstins(tabName) { return mbDedupe_(tabName, false); }

function mbDedupe_(tabName, dryRun) {
  var ss = SpreadsheetApp.openById(CFG.SHEET_ID);
  var tabs = tabName ? [tabName] : CFG.QUERIES.map(function (q) { return q.tab; });
  var report = [];

  tabs.forEach(function (tab) {
    var sh = ss.getSheetByName(tab);
    if (!sh) { Logger.log('✗ Tab "' + tab + '" not found'); return; }

    var headerRow = mbFindHeaderRow_(sh);
    var lastRow   = sh.getLastRow();
    if (lastRow <= headerRow) { Logger.log('"' + tab + '": no data rows'); return; }

    var headers = sh.getRange(headerRow, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
    var rG = mbResolveSheetCol_(headers, 'gstin', tab);
    if (rG.idx < 0) { Logger.log('✗ "' + tab + '": GSTIN column unresolved — skipped'); return; }

    var firstDataRow = headerRow + 1;
    var n = lastRow - headerRow;
    var vals = sh.getRange(firstDataRow, rG.idx + 1, n, 1).getValues();

    var seen = {}, dupRows = [];
    for (var i = 0; i < n; i++) {
      var k = mbGstinKey_(vals[i][0]);
      if (!k) continue;                       // never touch blank-key rows
      if (seen[k]) dupRows.push(firstDataRow + i); else seen[k] = true;
    }

    Logger.log('"' + tab + '": ' + Object.keys(seen).length + ' unique GSTINs, ' +
               dupRows.length + ' duplicate row(s)' + (dryRun ? ' (preview only)' : ''));
    if (dupRows.length && dupRows.length <= 60) Logger.log('  Rows: ' + dupRows.join(', '));

    if (!dryRun && dupRows.length) {
      // Bottom-up so earlier indices remain valid as rows are removed.
      for (var d = dupRows.length - 1; d >= 0; d--) sh.deleteRow(dupRows[d]);
      Logger.log('  ✓ Deleted ' + dupRows.length + ' duplicate row(s), kept the first of each');

      var rNo = mbResolveSheetCol_(headers, 'no', tab);
      if (rNo.idx >= 0) renumberNoColumn_(sh, rNo.idx, rG.idx, headerRow);
    }
    report.push({ tab: tab, unique: Object.keys(seen).length, duplicates: dupRows.length, removed: dryRun ? 0 : dupRows.length });
  });
  return report;
}

/** Fill in any blank Name cells for rows already in the sheet, matched by GSTIN.
    Only blank cells are written — an existing name is never overwritten. */
function backfillNames(tabName) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { Logger.log('Busy — try again.'); return; }
  try {
    var token = getSessionToken_();
    var ss    = SpreadsheetApp.openById(CFG.SHEET_ID);
    var qs    = CFG.QUERIES.filter(function (q) { return !tabName || q.tab === tabName; });

    qs.forEach(function (q) {
      var sh = ss.getSheetByName(q.tab);
      if (!sh) { Logger.log('✗ "' + q.tab + '" not found'); return; }

      var raw  = applyFilter_(fetchCardData_(token, q.id));
      var gIdx = mbFindIdx_(raw.cols, FIELD_MAP[q.gstinKey] || []);
      var nIdx = mbFindIdx_(raw.cols, FIELD_MAP[q.nameKey]  || []);
      if (nIdx < 0) for (var i = 0; i < raw.cols.length; i++) if (MB_FUZZY.name(norm_(raw.cols[i]))) { nIdx = i; break; }
      if (gIdx < 0 || nIdx < 0) { Logger.log('✗ "' + q.tab + '": GSTIN/Name not found in card output'); return; }

      var byGstin = {};
      raw.rows.forEach(function (row) {
        var k = mbGstinKey_(row[gIdx]);
        if (k && !byGstin[k]) byGstin[k] = String(row[nIdx] == null ? '' : row[nIdx]).trim();
      });

      var headerRow = mbFindHeaderRow_(sh);
      var headers   = sh.getRange(headerRow, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0];
      var rG = mbResolveSheetCol_(headers, 'gstin', q.tab);
      var rN = mbResolveSheetCol_(headers, 'name',  q.tab);
      if (rG.idx < 0 || rN.idx < 0) { Logger.log('✗ "' + q.tab + '": GSTIN/Name column unresolved'); return; }

      var n = sh.getLastRow() - headerRow;
      if (n < 1) return;
      var gCol = sh.getRange(headerRow + 1, rG.idx + 1, n, 1).getValues();
      var nCol = sh.getRange(headerRow + 1, rN.idx + 1, n, 1).getValues();

      var filled = 0, changed = false;
      for (var r = 0; r < n; r++) {
        if (String(nCol[r][0] == null ? '' : nCol[r][0]).trim() !== '') continue;   // keep existing
        var nm = byGstin[mbGstinKey_(gCol[r][0])];
        if (nm) { nCol[r][0] = nm; filled++; changed = true; }
      }
      if (changed) sh.getRange(headerRow + 1, rN.idx + 1, n, 1).setValues(nCol);
      Logger.log('"' + q.tab + '": filled ' + filled + ' blank name(s) in column ' + mbColLetter_(rN.idx));
    });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/* ─────────────────────────────────────────────────────────────────────────────
   LEGACY FULL-REFRESH (retained, not scheduled)
   ───────────────────────────────────────────────────────────────────────────
   The previous behaviour: clear rows 2..lastRow across the full header width and
   rewrite every column, including documents and vintage. It is destructive to
   anything maintained in the sheet by hand, which is why the scheduled sync no
   longer uses it. Kept only so the mapping/vintage helpers remain reachable for
   a deliberate, manual rebuild of a tab.
   ───────────────────────────────────────────────────────────────────────────*/
function legacyFullRefreshSync_() {
  var token = getSessionToken_();
  var ss    = SpreadsheetApp.openById(CFG.SHEET_ID);
  CFG.QUERIES.forEach(function (q) {
    var raw          = fetchCardData_(token, q.id);
    var gstinDateMap = buildEarliestDateMap_(raw);
    var filtered     = applyFilter_(raw);
    writeToSheet_(ss, q.tab, filtered, gstinDateMap);
    postProcessVintage_(ss, q.tab, filtered.rows.length, gstinDateMap);
    formatSheet_(ss, q.tab, filtered.rows.length);
  });
  Logger.log('══ Legacy full refresh complete ══');
}


// ─────────────────────────────────────────────────────────────────────────────
// DEBUG HELPER — run this to see what columns Metabase actually returns
// ─────────────────────────────────────────────────────────────────────────────

function debugColumns() {
  var token = getSessionToken_();
  CFG.QUERIES.forEach(function (q) {
    Logger.log('══ Card ' + q.id + ' (' + q.tab + ') ══');
    var raw = fetchCardData_(token, q.id);
    Logger.log('Total rows (unfiltered): ' + raw.rows.length);
    Logger.log('Columns (' + raw.cols.length + '):');
    raw.cols.forEach(function (c, i) { Logger.log('  [' + (i + 1) + '] ' + c); });
    // Sample first row
    if (raw.rows.length > 0) {
      Logger.log('First row sample:');
      raw.cols.forEach(function (c, i) { Logger.log('  ' + c + ' = ' + raw.rows[0][i]); });
    }
  });
}


// ─────────────────────────────────────────────────────────────────────────────
// SESSION MANAGEMENT  (cached in Script Properties — avoids login every minute)
// ─────────────────────────────────────────────────────────────────────────────

function getSessionToken_() {
  var props  = PropertiesService.getScriptProperties();
  var token  = props.getProperty('MB_TOKEN');
  var expStr = props.getProperty('MB_TOKEN_EXP');
  var exp    = expStr ? parseInt(expStr, 10) : 0;

  // Reuse cached token if it hasn't expired (we refresh every 6 h, sessions last 14 d)
  if (token && Date.now() < exp) {
    Logger.log('Using cached session (expires in ' +
      Math.round((exp - Date.now()) / 60000) + ' min)');
    return token;
  }

  Logger.log('Authenticating with ' + CFG.METABASE_URL + ' …');
  var resp = UrlFetchApp.fetch(CFG.METABASE_URL + '/api/session', {
    method            : 'post',
    contentType       : 'application/json',
    payload           : JSON.stringify({ username: CFG.METABASE_USER, password: CFG.METABASE_PASS }),
    muteHttpExceptions: true,
  });

  var code = resp.getResponseCode();
  if (code !== 200) {
    throw new Error(
      'Metabase login failed (HTTP ' + code + '). ' +
      'Check CFG.METABASE_USER / CFG.METABASE_PASS.\n' +
      resp.getContentText().substring(0, 300)
    );
  }

  token = JSON.parse(resp.getContentText()).id;
  props.setProperty('MB_TOKEN',     token);
  props.setProperty('MB_TOKEN_EXP', String(Date.now() + 6 * 3600 * 1000));
  Logger.log('Login OK — token cached for 6 h');
  return token;
}

// Force a fresh login on the next run (call this if you get 401 errors)
function clearSession() {
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty('MB_TOKEN');
  props.deleteProperty('MB_TOKEN_EXP');
  Logger.log('Session cleared — next run will re-authenticate.');
}


// ─────────────────────────────────────────────────────────────────────────────
// METABASE DATA FETCH
// Strategy: try the JSON export endpoint first (no row cap), fall back to the
// standard card query endpoint (2 000-row default cap).
// ─────────────────────────────────────────────────────────────────────────────

function fetchCardData_(token, cardId) {
  var headers = { 'X-Metabase-Session': token };

  // ── Strategy 1: JSON export — bypasses the 2 000-row display cap ──────────
  var exportResp = UrlFetchApp.fetch(
    CFG.METABASE_URL + '/api/card/' + cardId + '/query/json',
    {
      method            : 'post',
      contentType       : 'application/json',
      headers           : headers,
      payload           : JSON.stringify({ parameters: [] }),
      muteHttpExceptions: true,
    }
  );

  var code = exportResp.getResponseCode();
  Logger.log('  /query/json → HTTP ' + code);

  if (code === 200) {
    var body = exportResp.getContentText();
    // Export returns a JSON array of objects: [{"Col": val, ...}, ...]
    try {
      var arr = JSON.parse(body);
      if (Array.isArray(arr) && arr.length > 0) {
        var cols = Object.keys(arr[0]);
        var rows = arr.map(function (obj) { return cols.map(function (k) { return obj[k]; }); });
        Logger.log('  Export parsed: ' + rows.length + ' rows via /query/json');
        return { cols: cols, rows: rows };
      }
      if (Array.isArray(arr) && arr.length === 0) {
        Logger.log('  Export returned 0 rows — query may genuinely be empty');
        return { cols: [], rows: [] };
      }
    } catch (e) {
      Logger.log('  /query/json parse error: ' + e.message + ' — falling back');
    }
  }

  if (code === 401) {
    // Clear stale session so next run re-authenticates
    clearSession();
    throw new Error('Session expired (401). Cleared — will re-authenticate on next run.');
  }

  // ── Strategy 2: Standard card query endpoint ───────────────────────────────
  Logger.log('  Falling back to /query …');
  var queryResp = UrlFetchApp.fetch(
    CFG.METABASE_URL + '/api/card/' + cardId + '/query',
    {
      method            : 'post',
      contentType       : 'application/json',
      headers           : headers,
      payload           : JSON.stringify({ parameters: [], ignore_cache: false }),
      muteHttpExceptions: true,
    }
  );

  code = queryResp.getResponseCode();
  Logger.log('  /query → HTTP ' + code);

  if (code === 401) { clearSession(); throw new Error('Session expired (401). Will re-auth on next run.'); }
  if (code !== 200 && code !== 202) {
    throw new Error(
      'Card ' + cardId + ' query failed (HTTP ' + code + '):\n' +
      queryResp.getContentText().substring(0, 400)
    );
  }

  var qBody = JSON.parse(queryResp.getContentText());
  var data  = qBody.data || qBody;

  if (!data.cols || !data.rows) {
    Logger.log('  Unexpected response shape: ' + JSON.stringify(qBody).substring(0, 400));
    throw new Error('Unexpected response format from card ' + cardId + ' /query endpoint.');
  }

  var qCols = data.cols.map(function (c) { return String(c.display_name || c.name || '').trim(); });
  var qRows = data.rows;

  if (data.rows_truncated != null) {
    Logger.log('  ⚠ rows_truncated present — result may be capped at 2 000 rows');
  }

  Logger.log('  /query parsed: ' + qRows.length + ' rows');
  return { cols: qCols, rows: qRows };
}


// ─────────────────────────────────────────────────────────────────────────────
// ROW FILTER
// ─────────────────────────────────────────────────────────────────────────────

function applyFilter_(result) {
  var cols = result.cols;

  function findIdx(candidates) {
    for (var i = 0; i < cols.length; i++) {
      var n = norm_(cols[i]);
      for (var j = 0; j < candidates.length; j++) {
        if (n === norm_(candidates[j])) return i;
      }
    }
    return -1;
  }

  var vIdx = findIdx(['Vertical', 'Business Vertical', 'BusinessVertical',
                      'business_vertical', 'Biz Vertical']);
  var oIdx = findIdx(['Onboarding Status', 'OnboardingStatus', 'Onboarding_Status',
                      'onboarding_status', 'Status', 'Seller Status']);

  var wV = String(CFG.FILTER.VERTICAL || '').toLowerCase().trim();
  var wO = String(CFG.FILTER.ONBOARDING_STATUS || '').toLowerCase().trim();

  Logger.log('  Filter — Vertical col idx: ' + vIdx + (wO ? ', Onboarding col idx: ' + oIdx : ', Onboarding filter: off'));
  if (wV && vIdx < 0) Logger.log('  ⚠ Vertical column not found — run debugColumns() to see exact names');
  if (wO && oIdx < 0) Logger.log('  ⚠ Onboarding Status column not found — run debugColumns() to see exact names');

  var filtered = result.rows.filter(function (row) {
    // An empty CFG value disables that filter entirely.
    var vOk = !wV || vIdx < 0 || String(row[vIdx] == null ? '' : row[vIdx]).toLowerCase().trim() === wV;
    var oOk = !wO || oIdx < 0 || String(row[oIdx] == null ? '' : row[oIdx]).toLowerCase().trim() === wO;
    return vOk && oOk;
  });

  return { cols: cols, rows: filtered };
}


// ─────────────────────────────────────────────────────────────────────────────
// DOCUMENT FIELD TRANSFORM
// 1 / 2 / 3  →  "Received"
// 0           →  "Not Received"
// null / ''   →  "NA"   (column present but blank)
// col absent  →  "NA"   (handled in writeToSheet_)
// ─────────────────────────────────────────────────────────────────────────────

function transformDocValue_(v) {
  if (v === null || v === undefined || String(v).trim() === '') return 'NA';
  var s = String(v).trim().toLowerCase();
  // Already transformed
  if (s === 'received')     return 'Received';
  if (s === 'not received') return 'Not Received';
  if (s === 'na' || s === 'n/a' || s === '-') return 'NA';
  // Numeric
  var n = parseFloat(s);
  if (!isNaN(n)) {
    if (n >= 1) return 'Received';
    if (n === 0) return 'Not Received';
  }
  // Non-empty non-numeric text — treat as Received
  return 'Received';
}


// ─────────────────────────────────────────────────────────────────────────────
// VINTAGE CALCULATION
// parseDate_(v)             — any date representation → JS Date or null
// calcVintage_(dateVal)     — date → "XY YM" string (e.g., "2Y 3M", "0Y 6M")
// buildEarliestDateMap_(r)  — GSTIN → earliest Date across ALL verticals
// ─────────────────────────────────────────────────────────────────────────────

function parseDate_(dateVal) {
  if (dateVal == null) return null;
  if (dateVal instanceof Date) return isNaN(dateVal.getTime()) ? null : dateVal;
  var s = String(dateVal).trim();
  if (!s) return null;
  var d = new Date(s);
  if (!isNaN(d.getTime())) return d;
  var n = parseFloat(s);
  if (!isNaN(n) && n > 1000) {                     // Excel date serial
    d = new Date(Date.UTC(1899, 11, 30) + n * 864e5);
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
}

function calcVintage_(dateVal) {
  var d = parseDate_(dateVal);
  if (!d) return '';
  var now = new Date();
  if (d > now) return '';
  var yrs = now.getFullYear() - d.getFullYear();
  var mos = now.getMonth()   - d.getMonth();
  if (mos < 0) { yrs--; mos += 12; }
  if (yrs < 0) return '';
  return yrs + 'Y ' + mos + 'M';
}

/** Walk ALL rows from a Metabase result (before vertical filtering) and build
 *  a map of GSTIN → earliest onboarding date. Vendors who joined Recykal in
 *  another vertical before Open Marketplace get credit for their full tenure. */
function buildEarliestDateMap_(result) {
  var map = {};
  if (!result || !result.rows || result.rows.length === 0) return map;

  var CL = {};
  result.cols.forEach(function(n, i) { CL[norm_(n)] = i; });

  // Resolve GSTIN column
  var gstinIdx = -1;
  var gstinCandidates = (FIELD_MAP['sellergstin'] || []).concat(FIELD_MAP['buyergstin'] || []).concat(['GSTIN', 'GST']);
  for (var gi = 0; gi < gstinCandidates.length && gstinIdx < 0; gi++) {
    var gk = norm_(gstinCandidates[gi]);
    if (gk in CL) gstinIdx = CL[gk];
  }
  if (gstinIdx < 0) {
    Logger.log('  ⚠ buildEarliestDateMap_: GSTIN column not found — vintage will fall back to row-level date');
    return map;
  }

  // Resolve onboarding / registration date column
  var dateIdx = -1;
  var dateAliases = FIELD_MAP['dateofregistration'] || [];
  for (var di = 0; di < dateAliases.length && dateIdx < 0; di++) {
    var dk = norm_(dateAliases[di]);
    if (dk in CL) dateIdx = CL[dk];
  }
  if (dateIdx < 0) {
    Logger.log('  ⚠ buildEarliestDateMap_: Onboarding date column not found — vintage will fall back to row-level date');
    return map;
  }

  Logger.log('  buildEarliestDateMap_: GSTIN=[' + gstinIdx + '] "' + result.cols[gstinIdx] + '"  date=[' + dateIdx + '] "' + result.cols[dateIdx] + '"');

  // Track minimum date per GSTIN across all rows (all verticals)
  result.rows.forEach(function(row) {
    var gstin = String(row[gstinIdx] == null ? '' : row[gstinIdx]).trim().toUpperCase();
    if (!gstin) return;
    var d = parseDate_(row[dateIdx]);
    if (!d) return;
    if (!map[gstin] || d < map[gstin]) map[gstin] = d;
  });

  Logger.log('  buildEarliestDateMap_: ' + Object.keys(map).length + ' GSTINs indexed across all verticals');
  return map;
}

/** After writeToSheet_() has written everything else, read the sheet's own
 *  headers to locate the Vintage and date columns, then calculate and write
 *  vintage values directly. This is reliable regardless of Metabase column
 *  names because we read the date from the sheet (already written by sync). */
function postProcessVintage_(ss, tabName, numRows, gstinDateMap) {
  if (numRows <= 0) return;
  var sh = ss.getSheetByName(tabName);
  if (!sh) return;

  var lastCol = sh.getLastColumn();
  if (lastCol <= 0) return;

  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function(h) { return String(h).trim(); });

  var vintageCol = -1;   // 1-based sheet column numbers
  var gstinCol   = -1;
  var dateCol    = -1;

  var gstinAliases = (FIELD_MAP['sellergstin'] || [])
    .concat(FIELD_MAP['buyergstin'] || [])
    .concat(['GSTIN', 'GST']);
  var dateAliases = FIELD_MAP['dateofregistration'] || [];

  headers.forEach(function(h, i) {
    var hn = norm_(h);
    if (vintageCol < 0 && hn === 'vintagewithrecykal')  vintageCol = i + 1;
    if (gstinCol < 0) {
      for (var g = 0; g < gstinAliases.length; g++) {
        if (norm_(gstinAliases[g]) === hn) { gstinCol = i + 1; break; }
      }
    }
    if (dateCol < 0) {
      for (var d = 0; d < dateAliases.length; d++) {
        if (norm_(dateAliases[d]) === hn) { dateCol = i + 1; break; }
      }
    }
  });

  if (vintageCol < 0) {
    Logger.log('  ⚠ postProcessVintage_: "Vintage with Recykal" header not found in "' + tabName + '" — all sheet headers: ' + headers.join(' | '));
    return;
  }

  Logger.log('  postProcessVintage_: vintage=col' + vintageCol
    + '  gstin=col' + (gstinCol > 0 ? gstinCol + ' "' + headers[gstinCol-1] + '"' : '-1 (not found)')
    + '  date=col'  + (dateCol  > 0 ? dateCol  + ' "' + headers[dateCol -1] + '"' : '-1 (not found)')
    + '  gstinMap=' + Object.keys(gstinDateMap || {}).length + ' entries');

  var data = sh.getRange(2, 1, numRows, lastCol).getValues();

  var vintageVals = data.map(function(row) {
    // 1. Earliest cross-vertical date from GSTIN map (built before vertical filter)
    if (gstinCol > 0 && gstinDateMap) {
      var gstin = String(row[gstinCol - 1] == null ? '' : row[gstinCol - 1]).trim().toUpperCase();
      if (gstin && gstinDateMap[gstin]) {
        var v1 = calcVintage_(gstinDateMap[gstin]);
        if (v1) return [v1];
      }
    }
    // 2. Date already written to the sheet's own date column
    if (dateCol > 0) {
      var v2 = calcVintage_(row[dateCol - 1]);
      if (v2) return [v2];
    }
    return [''];
  });

  sh.getRange(2, vintageCol, numRows, 1).setValues(vintageVals);
  Logger.log('  postProcessVintage_: ✓ wrote ' + numRows + ' vintage values to col ' + vintageCol + ' ("' + headers[vintageCol-1] + '")');
}


// ─────────────────────────────────────────────────────────────────────────────
// SHEET WRITER
// ─────────────────────────────────────────────────────────────────────────────

function writeToSheet_(ss, tabName, result, gstinDateMap) {
  var sh = ss.getSheetByName(tabName);
  if (!sh) throw new Error('Tab "' + tabName + '" not found in spreadsheet ' + CFG.SHEET_ID);

  // ── Read existing headers from row 1 (never overwrite them) ───────────────
  var lastCol = sh.getLastColumn();
  var headers;
  if (lastCol > 0) {
    headers = sh.getRange(1, 1, 1, lastCol).getValues()[0]
      .map(function (h) { return String(h).trim(); });
    while (headers.length > 0 && headers[headers.length - 1] === '') headers.pop();
  }
  if (!headers || headers.length === 0) {
    Logger.log('  Row 1 empty — writing default COLUMN_ORDER as headers');
    headers = COLUMN_ORDER.slice();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  Logger.log('  Sheet headers (' + headers.length + '): ' + headers.join(' | '));

  // ── Build lookup: normalised Metabase col name → index in result.cols ─────
  var colLookup = {};
  result.cols.forEach(function (name, i) { colLookup[norm_(name)] = i; });

  // ── Map each sheet header to a result column index (-1 = not found) ───────
  // Also track the matched Metabase column name so we can apply transforms.
  // Three-step resolution per header:
  //   1. FIELD_MAP explicit aliases (handles "Seller Business Name" → "Entity Name" etc.)
  //   2. Direct normalised name match
  //   3. Prefix-strip fuzzy ("seller"/"buyer" dropped from either side)
  var mapping        = [];   // result col index for each header (-1 if unmatched)
  var matchedSrcName = [];   // Metabase column name that was matched ('' if none)

  // The tab decides which side's aliases apply — not the header wording. Both
  // tabs get seeded from COLUMN_ORDER, which is seller-worded, so the Buyers tab
  // reads "Seller Business Name" and would otherwise never try the buyer aliases.
  var isBuyerTab = norm_(tabName).indexOf('buyer') >= 0;

  headers.forEach(function (h) {
    var key     = norm_(h);
    var idx     = -1;
    var srcName = '';

    // Step 1 — FIELD_MAP, tab-appropriate key first
    var keys = [key];
    if (isBuyerTab && key.indexOf('seller') === 0) keys.unshift('buyer'  + key.slice(6));
    if (!isBuyerTab && key.indexOf('buyer') === 0) keys.unshift('seller' + key.slice(5));

    for (var ki = 0; ki < keys.length && idx < 0; ki++) {
      var aliases = FIELD_MAP[keys[ki]];
      if (!aliases) continue;
      for (var fi = 0; fi < aliases.length; fi++) {
        var an = norm_(aliases[fi]);
        if (an in colLookup) {
          idx     = colLookup[an];
          srcName = result.cols[idx];
          break;
        }
      }
    }

    // Step 2 — Direct norm match
    if (idx < 0 && key in colLookup) {
      idx     = colLookup[key];
      srcName = result.cols[idx];
    }

    // Step 3 — Prefix-strip fuzzy
    if (idx < 0) {
      var stripped = key.replace(/^(seller|buyer)/, '');
      for (var k in colLookup) {
        var ks = k.replace(/^(seller|buyer)/, '');
        if (k === stripped || ks === key || ks === stripped) {
          idx     = colLookup[k];
          srcName = result.cols[idx];
          break;
        }
      }
    }

    // Step 4 — Substring fallback: sheet-header norm is a substring of a Metabase
    //          col norm, or vice versa. Only for keys ≥ 5 chars (avoids false
    //          positives on short keys like 'no', 'id', 'gstin').
    if (idx < 0 && key.length >= 5) {
      var bestK = null, bestLen = 0;
      for (var k in colLookup) {
        if (k.length < 4) continue;
        if (k.indexOf(key) >= 0 || key.indexOf(k) >= 0) {
          if (k.length > bestLen) { bestLen = k.length; bestK = k; }
        }
      }
      if (bestK !== null) {
        idx     = colLookup[bestK];
        srcName = result.cols[idx];
      }
    }

    mapping.push(idx);
    matchedSrcName.push(srcName);
  });

  // ── Audit: log what matched and what didn't ────────────────────────────────
  Logger.log('  ── Column mapping (' + tabName + ') ──');
  Logger.log('  Metabase columns available: ' + result.cols.join(' | '));
  headers.forEach(function (h, c) {
    var isDoc = norm_(h) in DOC_COL_NORMS;
    var status = mapping[c] >= 0
      ? '✓ → [' + mapping[c] + '] "' + matchedSrcName[c] + '"'
      : (isDoc ? '◌ → NA (doc col not in query)' : '✗ → NO MATCH (will be blank)');
    Logger.log('    [' + (c + 1) + '] ' + h + '  ' + status);
  });

  // The business-name column failing to resolve is the one miss that looks like
  // a working sync — every row writes, the name is just blank. Call it out with
  // the query's actual column names so the alias can be added in one pass.
  var nameCol = -1;
  headers.forEach(function (h, c) { if (norm_(h).indexOf('businessname') >= 0) nameCol = c; });
  if (nameCol >= 0 && mapping[nameCol] < 0) {
    Logger.log(
      '  ⚠ BUSINESS NAME UNRESOLVED on "' + tabName + '" — that column will be blank.\n' +
      '    Sheet header : "' + headers[nameCol] + '"\n' +
      '    Query returns: ' + result.cols.join(' | ') + '\n' +
      '    Fix: add the correct name to FIELD_MAP[\'' +
        (isBuyerTab ? 'buyerbusinessname' : 'sellerbusinessname') + '\'].'
    );
  }

  // ── Resolve GSTIN and onboarding-date source columns ─────────────────────
  // gstinSrcIdx    — joins each filtered row to gstinDateMap for cross-vertical
  //                  earliest-date lookup (built from ALL rows in syncMetabaseToSheet).
  // onboardDateSrcIdx — per-row fallback when the GSTIN map has no entry.
  var gstinSrcIdx = -1;
  var _gstinCandidates = (FIELD_MAP['sellergstin'] || []).concat(FIELD_MAP['buyergstin'] || []).concat(['GSTIN', 'GST']);
  for (var _gi = 0; _gi < _gstinCandidates.length && gstinSrcIdx < 0; _gi++) {
    var _gn = norm_(_gstinCandidates[_gi]);
    if (_gn in colLookup) gstinSrcIdx = colLookup[_gn];
  }
  var onboardDateSrcIdx = -1;
  var _dateAliases = FIELD_MAP['dateofregistration'] || [];
  for (var _da = 0; _da < _dateAliases.length && onboardDateSrcIdx < 0; _da++) {
    var _an = norm_(_dateAliases[_da]);
    if (_an in colLookup) onboardDateSrcIdx = colLookup[_an];
  }
  Logger.log('  GSTIN col: ' + (gstinSrcIdx >= 0 ? '[' + gstinSrcIdx + '] "' + result.cols[gstinSrcIdx] + '"' : '⚠ not found'));
  Logger.log('  Onboarding-date col: ' + (onboardDateSrcIdx >= 0
    ? '[' + onboardDateSrcIdx + '] "' + result.cols[onboardDateSrcIdx] + '"'
    : '⚠ not found — vintage will fall back to Metabase value'));

  // ── Build output rows ──────────────────────────────────────────────────────
  var outRows = result.rows.map(function (srcRow, rowIdx) {
    return headers.map(function (h, c) {
      var hNorm  = norm_(h);
      var isDoc  = hNorm in DOC_COL_NORMS;

      // "No." column — always auto-filled with a sequential 1-based serial number
      if (hNorm === 'no') return rowIdx + 1;

      // "Vintage with Recykal" — MUST come before the mapping[c] < 0 guard because
      // Metabase does not return this column; mapping[c] will be -1 and the guard
      // would return '' before any calculation runs.
      // Priority: (1) earliest date across all verticals from gstinDateMap,
      //           (2) date in this filtered row, (3) Metabase's own value, (4) ''.
      if (hNorm === 'vintagewithrecykal') {
        if (gstinSrcIdx >= 0 && gstinDateMap) {
          var _g = String(srcRow[gstinSrcIdx] == null ? '' : srcRow[gstinSrcIdx]).trim().toUpperCase();
          if (_g && gstinDateMap[_g]) {
            var _vc = calcVintage_(gstinDateMap[_g]);
            if (_vc) return _vc;
          }
        }
        if (onboardDateSrcIdx >= 0) {
          var _vf = calcVintage_(srcRow[onboardDateSrcIdx]);
          if (_vf) return _vf;
        }
        if (mapping[c] >= 0) { var _vm = srcRow[mapping[c]]; return _vm == null ? '' : _vm; }
        return '';
      }

      if (mapping[c] < 0) {
        return isDoc ? 'NA' : '';   // doc columns absent from query → NA
      }
      var v = srcRow[mapping[c]];
      if (isDoc) return transformDocValue_(v);
      // "Region" — always derived from the State value via North/South/NA logic
      if (hNorm === 'region') return deriveRegion_(v);
      return v == null ? '' : v;
    });
  });

  // ── Clear existing data rows (2 onwards), keep header ─────────────────────
  var lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, headers.length).clearContent();

  // ── Write ──────────────────────────────────────────────────────────────────
  if (outRows.length > 0) {
    sh.getRange(2, 1, outRows.length, headers.length).setValues(outRows);
  }

  // ── Validation summary ─────────────────────────────────────────────────────
  var matched   = mapping.filter(function (i) { return i >= 0; }).length;
  var unmatched = mapping.filter(function (i) { return i < 0;  }).length;
  Logger.log('  ✓ Rows written   : ' + outRows.length);
  Logger.log('  ✓ Cols matched   : ' + matched + ' / ' + headers.length);
  Logger.log('  ✓ Cols unmatched : ' + unmatched + ' (doc cols → NA, others → blank)');
  Logger.log('  ✓ Sheet total    : ' + (outRows.length + 1) + ' rows (incl. header)');
}


// ─────────────────────────────────────────────────────────────────────────────
// SHEET FORMATTER
// Applies after every write: styled header row, auto-sized columns, borders.
// Only touches formatting — never overwrites values.
// ─────────────────────────────────────────────────────────────────────────────

function formatSheet_(ss, tabName, dataRowCount) {
  var sh = ss.getSheetByName(tabName);
  if (!sh) return;

  var lastCol = sh.getLastColumn();
  if (lastCol < 1) return;

  // ── Header row (row 1) ────────────────────────────────────────────────────
  var headerRange = sh.getRange(1, 1, 1, lastCol);
  headerRange
    .setBackground('#1a3c5e')          // dark navy — matches screenshot
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(10)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);

  sh.setRowHeight(1, 40);             // taller header row

  // ── Freeze header so it stays visible while scrolling ────────────────────
  sh.setFrozenRows(1);

  // ── Data rows ─────────────────────────────────────────────────────────────
  if (dataRowCount > 0) {
    var dataRange = sh.getRange(2, 1, dataRowCount, lastCol);
    dataRange
      .setFontColor('#000000')
      .setFontSize(10)
      .setHorizontalAlignment('left')
      .setVerticalAlignment('middle')
      .setBackground('#ffffff');

    // Alternating row shading
    for (var r = 0; r < dataRowCount; r++) {
      var bg = (r % 2 === 0) ? '#ffffff' : '#f5f8fc';
      sh.getRange(r + 2, 1, 1, lastCol).setBackground(bg);
    }

    sh.setRowHeightsForced(2, dataRowCount, 22);
  }

  // ── Borders on full data range (header + data) ────────────────────────────
  var totalRows = 1 + dataRowCount;
  if (totalRows > 1) {
    sh.getRange(1, 1, totalRows, lastCol)
      .setBorder(
        true, true, true, true, true, true,
        '#c0cfe0',
        SpreadsheetApp.BorderStyle.SOLID
      );
  }

  // ── Auto-resize all columns to fit content ────────────────────────────────
  for (var c = 1; c <= lastCol; c++) {
    sh.autoResizeColumn(c);
    // Cap column width so no column becomes excessively wide
    var w = sh.getColumnWidth(c);
    if (w > 220) sh.setColumnWidth(c, 220);
    if (w < 60)  sh.setColumnWidth(c, 60);
  }

  Logger.log('  Formatting applied to "' + tabName + '"');
}


// ─────────────────────────────────────────────────────────────────────────────
// AUTO-SYNC TRIGGER MANAGEMENT
// ─────────────────────────────────────────────────────────────────────────────

function createAutoSync() {
  deleteAutoSync();
  ScriptApp.newTrigger('syncMetabaseToSheet')
    .timeBased()
    .everyMinutes(CFG.SYNC_EVERY_MINS)
    .create();
  Logger.log('Auto-sync ON: every ' + CFG.SYNC_EVERY_MINS + ' min. Run deleteAutoSync() to stop.');
}

function deleteAutoSync() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncMetabaseToSheet') {
      ScriptApp.deleteTrigger(t);
      removed++;
    }
  });
  Logger.log(removed ? 'Auto-sync OFF (' + removed + ' trigger(s) removed).' : 'No trigger was active.');
}


// ─────────────────────────────────────────────────────────────────────────────
// UTILITY
// ─────────────────────────────────────────────────────────────────────────────

// State-to-region lookup tables (keyed on lowercase trimmed state name)
var NORTH_STATES = {
  'delhi': true, 'haryana': true, 'himachal pradesh': true,
  'jammu & kashmir': true, 'jammu and kashmir': true, 'j&k': true,
  'ladakh': true, 'punjab': true, 'rajasthan': true,
  'uttar pradesh': true, 'up': true,
  'uttarakhand': true, 'uttaranchal': true, 'chandigarh': true,
};
var SOUTH_STATES = {
  'andhra pradesh': true, 'ap': true,
  'karnataka': true, 'kerala': true,
  'tamil nadu': true, 'tamilnadu': true, 'tn': true,
  'telangana': true, 'puducherry': true, 'pondicherry': true,
  'lakshadweep': true,
  'andaman & nicobar islands': true, 'andaman and nicobar islands': true,
  'andaman & nicobar': true, 'andaman and nicobar': true,
};

// Derive Region from a State name, or pass through an already-computed region.
// Accepts "North"/"South"/"East"/"West" directly if Meta returns them.
// Returns "North", "South", or "NA" for unrecognised state names.
function deriveRegion_(stateValue) {
  if (stateValue == null || String(stateValue).trim() === '') return 'NA';
  var s = String(stateValue).trim().toLowerCase();
  // Already a region label — pass through
  if (s === 'north') return 'North';
  if (s === 'south') return 'South';
  if (s === 'east')  return 'East';
  if (s === 'west')  return 'West';
  if (NORTH_STATES[s]) return 'North';
  if (SOUTH_STATES[s]) return 'South';
  return 'NA';
}

// Normalise a string for fuzzy column matching
function norm_(s) {
  return String(s).toLowerCase().replace(/[\s_\-\/(),\.]+/g, '').trim();
}
