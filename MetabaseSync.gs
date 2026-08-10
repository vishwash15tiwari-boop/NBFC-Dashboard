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
  /* tabAliases covers singular/plural and casing so an existing tab is always
     found. A tab is only ever created when none of the aliases match. */
  QUERIES: [
    { id: 5712, tab: 'Sellers', tabAliases: ['Sellers', 'Seller', 'Seller Data', 'Sellers Data'],
      gstinKey: 'sellergstin', nameKey: 'sellerbusinessname' },
    { id: 5711, tab: 'Buyers',  tabAliases: ['Buyer', 'Buyers', 'Buyer Data', 'Buyers Data'],
      gstinKey: 'buyergstin',  nameKey: 'buyerbusinessname'  },
  ],

  /* Authoritative sheet layout, 1-based and confirmed: A = No., B = GSTIN,
     C = Name. These positions are used directly — header text is only read to
     cross-check and warn, never to decide. Change a number here to relocate a
     field; set one to 0 to disable writing it. */
  COLUMNS: { no: 1, gstin: 2, name: 3 },

  /* Every run repairs the tab before it appends, so the sync is self-healing
     and no separate clean-up call is required. */
  AUTO_DEDUPE: true,
  AUTO_BACKFILL_NAMES: true,

  /* A record must satisfy EVERY rule below to be synced.
     `aliases` locate the column in the Metabase output; `values` are the
     accepted cell values, compared case- and punctuation-insensitively — so
     "Complete" and "Completed" both pass without editing logic.
     Set `values: []` to switch a rule off. */
  FILTER: {
    vertical: {
      label:   'Business Vertical',
      aliases: ['Vertical', 'Business Vertical', 'BusinessVertical', 'business_vertical',
                'Biz Vertical', 'Vertical Name'],
      values:  ['Open Marketplace'],
    },
    onboarding: {
      label:   'Onboarded Status',
      aliases: ['Onboarded Status', 'Onboarding Status', 'Onboard Status', 'Onboarding State',
                'Onboarded', 'OnboardingStatus', 'OnboardedStatus', 'Onboarding_Status',
                'Onboarded_Status', 'Seller Status', 'Buyer Status', 'Status'],
      values:  ['Complete', 'Completed'],
    },
  },

  /* When a filter is configured but its column is absent from the query output,
     abort that tab instead of syncing unfiltered. Appending records that should
     have been excluded is worse than appending none, and the log names the
     columns that were available so the alias list can be corrected. */
  STRICT_FILTER: true,

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
        if (filtered.filterError) {
          // Nothing is written when a configured filter could not be applied.
          Logger.log('✗ "' + q.tab + '" skipped — ' + filtered.filterError);
          return;
        }
        Logger.log('After filters: ' + filtered.rows.length + ' of ' + raw.rows.length + ' rows');
        if (filtered.rows.length === 0) {
          Logger.log('⚠ Zero rows after filter — run debugSheetColumns() and check CFG.FILTER values');
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

/**
 * Find the tab for a query. Tries the configured name, then every alias, then a
 * normalised comparison against all existing tabs. Returns null when nothing
 * matches, so the caller decides whether creating one is appropriate.
 */
function mbResolveTab_(ss, q) {
  var wanted = [q.tab].concat(q.tabAliases || []);
  for (var i = 0; i < wanted.length; i++) {
    var sh = ss.getSheetByName(wanted[i]);
    if (sh) return sh;
  }
  var all = ss.getSheets();
  for (var w = 0; w < wanted.length; w++) {
    var target = norm_(wanted[w]);
    for (var s = 0; s < all.length; s++) if (norm_(all[s].getName()) === target) return all[s];
  }
  return null;
}

/**
 * Last-resort detection of the name column in a query result.
 * A business-name column is text and nearly unique per row, whereas Vertical,
 * Status, State and Type repeat across rows — so the column with the highest
 * ratio of distinct values wins. Structural columns are excluded outright.
 * @return {number} column index, or -1.
 */
function mbGuessNameCol_(cols, rows, skipIdx) {
  if (!rows.length) return -1;
  var sample = rows.slice(0, 400);
  var best = -1, bestRatio = 0;

  for (var c = 0; c < cols.length; c++) {
    if (c === skipIdx) continue;
    var n = norm_(cols[c]);
    if (/gst|vertical|status|state|region|type|date|pan|mobile|phone|email|id$|^id|code|pincode|amount|limit|count|score|rating/.test(n)) continue;

    var distinct = {}, textRows = 0;
    for (var r = 0; r < sample.length; r++) {
      var v = sample[r][c];
      if (v == null || typeof v === 'number' || v instanceof Date) continue;
      var s = String(v).trim();
      if (s.length < 3) continue;
      if (/^\d+$/.test(s)) continue;                       // pure numbers are not names
      if (/^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]/i.test(s)) continue;   // looks like a GSTIN
      textRows++; distinct[s.toLowerCase()] = true;
    }
    if (textRows < Math.min(5, sample.length)) continue;
    var ratio = Object.keys(distinct).length / textRows;
    if (ratio > bestRatio) { bestRatio = ratio; best = c; }
  }
  // Names are near-unique; anything below this is a category column.
  return bestRatio >= 0.5 ? best : -1;
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
 * Bring one tab fully up to date, in a single pass:
 *   1. remove duplicate GSTIN rows (keeping the first, data-bearing one)
 *   2. append records whose GSTIN is not present
 *   3. fill blank Name cells for rows already there
 *   4. renumber "No."
 * Column positions come from CFG.COLUMNS and are used directly, so nothing
 * depends on guessing header text. Only those columns are ever written.
 * @return {{added:number, existing:number, blank:number, removed:number, named:number}}
 */
function appendNewEntities_(ss, q, result) {
  var out = { added: 0, existing: 0, blank: 0, removed: 0, named: 0 };

  // ── Locate GSTIN + Name in the Metabase output ───────────────────────────
  var gIdx = mbFindIdx_(result.cols, FIELD_MAP[q.gstinKey] || []);
  if (gIdx < 0) for (var gi = 0; gi < result.cols.length; gi++) {
    if (MB_FUZZY.gstin(norm_(result.cols[gi]))) { gIdx = gi; break; }
  }
  if (gIdx < 0) {
    Logger.log('✗ No GSTIN column in query output — nothing done. Columns: ' + result.cols.join(' | '));
    return out;
  }

  var nIdx = mbFindIdx_(result.cols, FIELD_MAP[q.nameKey] || []);
  var how  = 'alias';
  if (nIdx < 0) {
    for (var ni = 0; ni < result.cols.length; ni++) {
      if (MB_FUZZY.name(norm_(result.cols[ni]))) { nIdx = ni; how = 'fuzzy'; break; }
    }
  }
  if (nIdx < 0) {                       // neither header matched — infer from the data
    nIdx = mbGuessNameCol_(result.cols, result.rows, gIdx);
    how  = 'value-shape';
  }
  Logger.log('  Metabase → GSTIN: "' + result.cols[gIdx] + '"');
  if (nIdx >= 0) {
    Logger.log('  Metabase → Name : "' + result.cols[nIdx] + '" (matched by ' + how + ')');
  } else {
    Logger.log('  ⚠ Name column NOT found. Columns returned: ' + result.cols.join(' | '));
    Logger.log('    → add the correct header to FIELD_MAP.' + q.nameKey);
  }

  // ── Resolve the destination tab (never silently create a duplicate tab) ──
  var sh = mbResolveTab_(ss, q);
  if (!sh) {
    sh = ss.insertSheet(q.tab);
    sh.getRange(1, 1, 1, 3).setValues([['No.', 'GSTIN',
      q.gstinKey === 'buyergstin' ? 'Buyer Business Name' : 'Seller Business Name']]);
    sh.setFrozenRows(1);
    Logger.log('  ⚠ No tab matched ' + JSON.stringify([q.tab].concat(q.tabAliases || [])) +
               ' — created "' + q.tab + '"');
  }
  Logger.log('  Tab: "' + sh.getName() + '"');

  // ── Fixed column positions ───────────────────────────────────────────────
  var cNo    = (CFG.COLUMNS.no    || 0) - 1;
  var cGstin = (CFG.COLUMNS.gstin || 0) - 1;
  var cName  = (CFG.COLUMNS.name  || 0) - 1;
  if (cGstin < 0) { Logger.log('✗ CFG.COLUMNS.gstin is not set — nothing done.'); return out; }

  var headerRow = mbFindHeaderRow_(sh);
  var lastCol   = Math.max(sh.getLastColumn(), cName + 1, cGstin + 1, cNo + 1);
  var headers   = sh.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  Logger.log('  Header row ' + headerRow + ' → GSTIN=' + mbColLetter_(cGstin) +
             ' ("' + (headers[cGstin] || '') + '"), Name=' + mbColLetter_(cName) +
             ' ("' + (headers[cName] || '') + '"), No.=' + mbColLetter_(cNo) +
             ' ("' + (headers[cNo] || '') + '")');
  // Cross-check only: a surprising header is reported, never acted on.
  if (headers[cGstin] && !MB_FUZZY.gstin(norm_(headers[cGstin])) &&
      mbFindIdx_([headers[cGstin]], CFG.SHEET_COLS.gstin) < 0) {
    Logger.log('  ⚠ Column ' + mbColLetter_(cGstin) + ' is "' + headers[cGstin] +
               '", which does not look like GSTIN — check CFG.COLUMNS.');
  }

  // ── 1. Remove duplicate GSTIN rows already in the tab ────────────────────
  if (CFG.AUTO_DEDUPE) out.removed = mbDedupeSheet_(sh, cGstin, headerRow, false);

  // ── Existing GSTIN → row number, from one read of the key column ─────────
  var firstDataRow = headerRow + 1;
  var dataRows = Math.max(0, sh.getLastRow() - headerRow);
  var gstinCol = dataRows > 0 ? sh.getRange(firstDataRow, cGstin + 1, dataRows, 1).getValues() : [];
  var rowOf = {};
  for (var r = 0; r < gstinCol.length; r++) {
    var k = mbGstinKey_(gstinCol[r][0]);
    if (k && rowOf[k] == null) rowOf[k] = r;          // 0-based offset from firstDataRow
  }
  Logger.log('  Existing: ' + dataRows + ' row(s), ' + Object.keys(rowOf).length + ' unique GSTIN(s)');

  // ── Index the fetched records by GSTIN ───────────────────────────────────
  var fetched = {}, order = [];
  result.rows.forEach(function (row) {
    var key = mbGstinKey_(row[gIdx]);
    if (!key) { out.blank++; return; }
    if (fetched[key]) return;                          // duplicate inside this fetch
    fetched[key] = nIdx >= 0 ? String(row[nIdx] == null ? '' : row[nIdx]).trim() : '';
    order.push(key);
  });

  // ── 2. Append the genuinely new ones ─────────────────────────────────────
  var newKeys = order.filter(function (k) { return rowOf[k] == null; });
  out.existing = order.length - newKeys.length;

  if (newKeys.length) {
    var startRow = sh.getLastRow() + 1;
    var needed   = startRow + newKeys.length - 1;
    if (needed > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), needed - sh.getMaxRows());

    sh.getRange(startRow, cGstin + 1, newKeys.length, 1)
      .setValues(newKeys.map(function (k) { return [k]; }));
    if (cName >= 0) {
      sh.getRange(startRow, cName + 1, newKeys.length, 1)
        .setValues(newKeys.map(function (k) { return [fetched[k]]; }));
    }
    out.added = newKeys.length;
    Logger.log('  ✓ Appended ' + newKeys.length + ' new record(s) from row ' + startRow);
  } else {
    Logger.log('  ✓ No new records');
  }

  // ── 3. Backfill blank names on rows that were already present ────────────
  if (CFG.AUTO_BACKFILL_NAMES && cName >= 0 && nIdx >= 0) {
    var total = Math.max(0, sh.getLastRow() - headerRow);
    if (total > 0) {
      var gAll = sh.getRange(firstDataRow, cGstin + 1, total, 1).getValues();
      var nAll = sh.getRange(firstDataRow, cName  + 1, total, 1).getValues();
      var touched = false;
      for (var i = 0; i < total; i++) {
        if (String(nAll[i][0] == null ? '' : nAll[i][0]).trim() !== '') continue;  // keep existing
        var nm = fetched[mbGstinKey_(gAll[i][0])];
        if (nm) { nAll[i][0] = nm; out.named++; touched = true; }
      }
      if (touched) {
        sh.getRange(firstDataRow, cName + 1, total, 1).setValues(nAll);
        Logger.log('  ✓ Filled ' + out.named + ' blank name(s) in column ' + mbColLetter_(cName));
      }
    }
  }

  // ── 4. Renumber, even when nothing was appended ──────────────────────────
  if (cNo >= 0) renumberNoColumn_(sh, cNo, cGstin, headerRow);

  Logger.log('  Summary: +' + out.added + ' added, ' + out.existing + ' already present, ' +
             out.removed + ' duplicate(s) removed, ' + out.named + ' name(s) filled, ' +
             out.blank + ' without GSTIN');
  return out;
}

/**
 * Delete duplicate GSTIN rows in a sheet, keeping the first occurrence of each.
 * Rows are removed bottom-up so indices stay valid.
 * @return {number} rows removed (or that would be, when dryRun).
 */
function mbDedupeSheet_(sh, cGstin, headerRow, dryRun) {
  var firstDataRow = headerRow + 1;
  var n = sh.getLastRow() - headerRow;
  if (n < 1) return 0;

  var vals = sh.getRange(firstDataRow, cGstin + 1, n, 1).getValues();
  var seen = {}, dup = [];
  for (var i = 0; i < n; i++) {
    var k = mbGstinKey_(vals[i][0]);
    if (!k) continue;                                  // blank-key rows are left alone
    if (seen[k]) dup.push(firstDataRow + i); else seen[k] = true;
  }
  if (!dup.length) return 0;

  Logger.log('  ' + (dryRun ? 'Would remove ' : 'Removing ') + dup.length +
             ' duplicate row(s)' + (dup.length <= 40 ? ': ' + dup.join(', ') : ''));
  if (!dryRun) for (var d = dup.length - 1; d >= 0; d--) sh.deleteRow(dup[d]);
  return dup.length;
}

/**
 * Renumber the "No." column 1..N across rows that have a GSTIN.
 * Written only when the current numbering is actually wrong, so a no-op sync
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
  Logger.log('Tabs present: ' + ss.getSheets().map(function (x) { return '"' + x.getName() + '"'; }).join(', '));
  Logger.log('CFG.COLUMNS → No.=' + mbColLetter_(CFG.COLUMNS.no - 1) +
             ', GSTIN=' + mbColLetter_(CFG.COLUMNS.gstin - 1) +
             ', Name=' + mbColLetter_(CFG.COLUMNS.name - 1));

  var cGstin = CFG.COLUMNS.gstin - 1, cName = CFG.COLUMNS.name - 1;

  CFG.QUERIES.forEach(function (q) {
    Logger.log('\n──── Query "' + q.tab + '" ────');
    var sh = mbResolveTab_(ss, q);
    if (!sh) {
      Logger.log('  ✗ No tab matches ' + JSON.stringify([q.tab].concat(q.tabAliases || [])));
      return;
    }
    Logger.log('  Resolved tab: "' + sh.getName() + '"');

    var headerRow = mbFindHeaderRow_(sh);
    var lastCol   = Math.max(1, sh.getLastColumn());
    var headers   = sh.getRange(headerRow, 1, 1, lastCol).getValues()[0];
    Logger.log('  Rows: ' + sh.getLastRow() + '  Cols: ' + lastCol + '  Header row: ' + headerRow);
    headers.forEach(function (h, i) {
      if (String(h).trim() !== '') Logger.log('    ' + mbColLetter_(i) + ': "' + h + '"');
    });

    var n = sh.getLastRow() - headerRow;
    if (n > 0) {
      var g = sh.getRange(headerRow + 1, cGstin + 1, n, 1).getValues();
      var nm = cName >= 0 ? sh.getRange(headerRow + 1, cName + 1, n, 1).getValues() : null;
      var seen = {}, dups = 0, blanks = 0, blankNames = 0;
      for (var i = 0; i < n; i++) {
        var k = mbGstinKey_(g[i][0]);
        if (!k) { blanks++; } else if (seen[k]) { dups++; } else { seen[k] = true; }
        if (nm && k && String(nm[i][0] == null ? '' : nm[i][0]).trim() === '') blankNames++;
      }
      Logger.log('  Column ' + mbColLetter_(cGstin) + ' → ' + Object.keys(seen).length +
                 ' unique GSTIN(s), ' + dups + ' duplicate row(s), ' + blanks + ' blank');
      Logger.log('  Column ' + mbColLetter_(cName) + ' → ' + blankNames + ' row(s) with a GSTIN but no name');
      Logger.log('  First 3 data rows: ' + JSON.stringify(
        sh.getRange(headerRow + 1, 1, Math.min(3, n), Math.min(4, lastCol)).getValues()));
    }
  });

  var token = getSessionToken_();
  CFG.QUERIES.forEach(function (q) {
    Logger.log('\n──── Metabase card ' + q.id + ' (' + q.tab + ') ────');
    try {
      var raw = fetchCardData_(token, q.id);
      Logger.log('  ' + raw.rows.length + ' rows returned. Columns with a sample value:');
      raw.cols.forEach(function (c, i) {
        var sample = raw.rows.length ? String(raw.rows[0][i] == null ? '' : raw.rows[0][i]).substring(0, 40) : '';
        Logger.log('    "' + c + '"  e.g. "' + sample + '"');
      });
      var g = mbFindIdx_(raw.cols, FIELD_MAP[q.gstinKey] || []);
      var nA = mbFindIdx_(raw.cols, FIELD_MAP[q.nameKey] || []);
      var nF = -1; for (var i = 0; i < raw.cols.length; i++) if (MB_FUZZY.name(norm_(raw.cols[i]))) { nF = i; break; }
      var nG = mbGuessNameCol_(raw.cols, raw.rows, g);
      Logger.log('  → GSTIN     : ' + (g  >= 0 ? '"' + raw.cols[g]  + '"' : 'UNRESOLVED'));
      Logger.log('  → Name alias: ' + (nA >= 0 ? '"' + raw.cols[nA] + '"' : 'no match'));
      Logger.log('  → Name fuzzy: ' + (nF >= 0 ? '"' + raw.cols[nF] + '"' : 'no match'));
      Logger.log('  → Name guess: ' + (nG >= 0 ? '"' + raw.cols[nG] + '"' : 'no match'));
      var f = applyFilter_(raw);
      Logger.log('  After filters: ' + f.rows.length + ' of ' + raw.rows.length + ' rows' +
                 (f.filterError ? '  ✗ ' + f.filterError : ''));
    } catch (e) { Logger.log('  ✗ ' + (e && e.message || e)); }
  });
  Logger.log('\n════ end ════');
}

/**
 * Dry run of the Open Marketplace + Onboarded Status rules. Writes NOTHING, so
 * it is safe against production at any time.
 *
 * debugSheetColumns() already reports how many rows survive the filter, but a
 * count alone cannot answer the question that actually matters — "is it really
 * excluding the records I expect, and on what grounds?". This names the excluded
 * vendors and the rule and cell value that rejected each one, so the filter can
 * be checked against Metabase row by row instead of taken on trust.
 *
 * Run before syncMetabaseToSheet() whenever the rules or the query change.
 */
function previewFilter() {
  var token = getSessionToken_();
  Logger.log('════ FILTER PREVIEW — nothing will be written ════');
  Object.keys(CFG.FILTER).forEach(function (k) {
    var r = CFG.FILTER[k];
    Logger.log('  Rule "' + (r && r.label || k) + '" → '
      + (r && r.values && r.values.length ? r.values.join(' / ') : '(off)'));
  });
  Logger.log('  STRICT_FILTER: ' + (CFG.STRICT_FILTER ? 'on' : 'off'));

  CFG.QUERIES.forEach(function (q) {
    Logger.log('');
    Logger.log('──── card ' + q.id + ' → "' + q.tab + '" ────');
    try {
      var raw = fetchCardData_(token, q.id);
      var res = applyFilter_(raw);

      if (res.filterError) {
        Logger.log('  ✗ WOULD ABORT this tab — ' + res.filterError);
        return;
      }

      var excluded = raw.rows.length - res.rows.length;
      Logger.log('  Fetched       : ' + raw.rows.length);
      Logger.log('  WOULD SYNC    : ' + res.rows.length);
      Logger.log('  WOULD EXCLUDE : ' + excluded);
      if (!excluded) return;

      /* Resolve each active rule's column again so a per-row reason can be
         attributed. applyFilter_ aggregates its rejections by value; this needs
         them per record, which is the point of the preview. */
      var gIdx  = mbFindIdx_(raw.cols, FIELD_MAP[q.gstinKey] || []);
      var nIdx  = mbFindIdx_(raw.cols, FIELD_MAP[q.nameKey]  || []);
      if (nIdx < 0) nIdx = mbGuessNameCol_(raw.cols, raw.rows, gIdx);

      var rules = [];
      Object.keys(CFG.FILTER).forEach(function (k) {
        var r = CFG.FILTER[k];
        if (!r || !r.values || !r.values.length) return;
        var idx = -1;
        (r.aliases || []).forEach(function (a) {
          if (idx >= 0) return;
          for (var i = 0; i < raw.cols.length; i++) if (norm_(raw.cols[i]) === norm_(a)) { idx = i; return; }
        });
        if (idx < 0) return;
        var acc = {};
        r.values.forEach(function (v) { acc[norm_(v)] = true; });
        rules.push({ idx: idx, acc: acc, label: r.label || k });
      });

      var CAP = 20, shown = 0;
      Logger.log('  Excluded records:');
      for (var i = 0; i < raw.rows.length && shown < CAP; i++) {
        var row = raw.rows[i], why = null;
        for (var j = 0; j < rules.length; j++) {
          if (!rules[j].acc[norm_(row[rules[j].idx])]) {
            var seen = String(row[rules[j].idx] == null ? '' : row[rules[j].idx]).trim() || '(blank)';
            why = rules[j].label + ' = "' + seen + '"';
            break;
          }
        }
        if (!why) continue;
        Logger.log('    ✗ ' + (nIdx >= 0 ? String(row[nIdx]) : '(row ' + (i + 1) + ')')
          + (gIdx >= 0 ? '  [' + row[gIdx] + ']' : '') + '  — ' + why);
        shown++;
      }
      if (excluded > shown) Logger.log('    … and ' + (excluded - shown) + ' more');
    } catch (e) {
      Logger.log('  ✗ ' + (e && e.message || e));
    }
  });

  Logger.log('');
  Logger.log('════ nothing was written — run syncMetabaseToSheet() to apply ════');
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
  var qs = CFG.QUERIES.filter(function (q) {
    return !tabName || norm_(q.tab) === norm_(tabName) ||
           (q.tabAliases || []).some(function (a) { return norm_(a) === norm_(tabName); });
  });
  if (!qs.length) { Logger.log('✗ No configured tab matches "' + tabName + '"'); return []; }

  var cGstin = (CFG.COLUMNS.gstin || 0) - 1;
  var cNo    = (CFG.COLUMNS.no    || 0) - 1;
  var report = [];

  qs.forEach(function (q) {
    var sh = mbResolveTab_(ss, q);
    if (!sh) { Logger.log('✗ Tab for "' + q.tab + '" not found'); return; }

    var headerRow = mbFindHeaderRow_(sh);
    if (sh.getLastRow() <= headerRow) { Logger.log('"' + sh.getName() + '": no data rows'); return; }

    var removed = mbDedupeSheet_(sh, cGstin, headerRow, dryRun);
    var n = sh.getLastRow() - headerRow;
    var vals = sh.getRange(headerRow + 1, cGstin + 1, n, 1).getValues();
    var uniq = {}; vals.forEach(function (r) { var k = mbGstinKey_(r[0]); if (k) uniq[k] = true; });

    Logger.log('"' + sh.getName() + '": ' + Object.keys(uniq).length + ' unique GSTIN(s), ' +
               removed + ' duplicate row(s)' + (dryRun ? ' (preview only — nothing changed)' : ' removed'));

    if (!dryRun && removed && cNo >= 0) renumberNoColumn_(sh, cNo, cGstin, headerRow);
    report.push({ tab: sh.getName(), unique: Object.keys(uniq).length,
                  duplicates: removed, removed: dryRun ? 0 : removed });
  });
  return report;
}

/** Fill in blank Name cells for rows already in the sheet, matched by GSTIN.
    Only blank cells are written — an existing name is never overwritten.
    The normal sync does this automatically; this is here for a one-off pass. */
function backfillNames(tabName) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) { Logger.log('Busy — try again.'); return; }
  try {
    var token = getSessionToken_();
    var ss    = SpreadsheetApp.openById(CFG.SHEET_ID);
    var cGstin = (CFG.COLUMNS.gstin || 0) - 1;
    var cName  = (CFG.COLUMNS.name  || 0) - 1;
    if (cGstin < 0 || cName < 0) { Logger.log('✗ CFG.COLUMNS.gstin/name not set'); return; }

    CFG.QUERIES.filter(function (q) {
      return !tabName || norm_(q.tab) === norm_(tabName) ||
             (q.tabAliases || []).some(function (a) { return norm_(a) === norm_(tabName); });
    }).forEach(function (q) {
      var sh = mbResolveTab_(ss, q);
      if (!sh) { Logger.log('✗ Tab for "' + q.tab + '" not found'); return; }

      var raw  = applyFilter_(fetchCardData_(token, q.id));
      var gIdx = mbFindIdx_(raw.cols, FIELD_MAP[q.gstinKey] || []);
      if (gIdx < 0) for (var gi = 0; gi < raw.cols.length; gi++) if (MB_FUZZY.gstin(norm_(raw.cols[gi]))) { gIdx = gi; break; }
      var nIdx = mbFindIdx_(raw.cols, FIELD_MAP[q.nameKey] || []);
      if (nIdx < 0) for (var ni = 0; ni < raw.cols.length; ni++) if (MB_FUZZY.name(norm_(raw.cols[ni]))) { nIdx = ni; break; }
      if (nIdx < 0) nIdx = mbGuessNameCol_(raw.cols, raw.rows, gIdx);
      if (gIdx < 0 || nIdx < 0) {
        Logger.log('✗ "' + sh.getName() + '": GSTIN/Name not found in card output — ' + raw.cols.join(' | '));
        return;
      }
      Logger.log('"' + sh.getName() + '": using Metabase name column "' + raw.cols[nIdx] + '"');

      var byGstin = {};
      raw.rows.forEach(function (row) {
        var k = mbGstinKey_(row[gIdx]);
        if (k && !byGstin[k]) byGstin[k] = String(row[nIdx] == null ? '' : row[nIdx]).trim();
      });

      var headerRow = mbFindHeaderRow_(sh);
      var n = sh.getLastRow() - headerRow;
      if (n < 1) { Logger.log('  no data rows'); return; }

      var gCol = sh.getRange(headerRow + 1, cGstin + 1, n, 1).getValues();
      var nCol = sh.getRange(headerRow + 1, cName  + 1, n, 1).getValues();
      var filled = 0;
      for (var r = 0; r < n; r++) {
        if (String(nCol[r][0] == null ? '' : nCol[r][0]).trim() !== '') continue;
        var nm = byGstin[mbGstinKey_(gCol[r][0])];
        if (nm) { nCol[r][0] = nm; filled++; }
      }
      if (filled) sh.getRange(headerRow + 1, cName + 1, n, 1).setValues(nCol);
      Logger.log('  ✓ filled ' + filled + ' blank name(s) in column ' + mbColLetter_(cName));
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

/**
 * Keep only rows satisfying every rule in CFG.FILTER.
 *
 * Each rule locates its column by alias and accepts any of its listed values,
 * compared through norm_() so casing, spacing and punctuation differences do not
 * matter ("Complete" and "Completed" are both configured explicitly).
 *
 * A rule whose column is missing is a hard failure under STRICT_FILTER: the
 * result is emptied and `filterError` is set, so the caller skips the tab rather
 * than appending records that should have been filtered out.
 *
 * @return {{cols:Array, rows:Array, filterError:(string|null)}}
 */
function applyFilter_(result) {
  var cols = result.cols;

  function findIdx(candidates) {
    for (var j = 0; j < candidates.length; j++) {        // alias priority
      var want = norm_(candidates[j]);
      for (var i = 0; i < cols.length; i++) if (norm_(cols[i]) === want) return i;
    }
    return -1;
  }

  var active = [], missing = [];
  Object.keys(CFG.FILTER).forEach(function (key) {
    var rule = CFG.FILTER[key];
    if (!rule || !rule.values || !rule.values.length) {
      Logger.log('  Filter "' + (rule && rule.label || key) + '": off');
      return;
    }
    var idx = findIdx(rule.aliases || []);
    if (idx < 0) { missing.push(rule.label || key); return; }

    var accepted = {};
    rule.values.forEach(function (v) { accepted[norm_(v)] = true; });
    active.push({ idx: idx, accepted: accepted, label: rule.label || key });
    Logger.log('  Filter "' + (rule.label || key) + '" → column "' + cols[idx] +
               '" must be one of: ' + rule.values.join(' / '));
  });

  if (missing.length) {
    var msg = 'Filter column(s) not found: ' + missing.join(', ') +
              '. Columns available: ' + cols.join(' | ');
    Logger.log('  ✗ ' + msg);
    if (CFG.STRICT_FILTER) {
      Logger.log('  ✗ STRICT_FILTER is on — skipping this tab rather than syncing unfiltered.');
      return { cols: cols, rows: [], filterError: msg };
    }
    Logger.log('  ⚠ STRICT_FILTER is off — continuing without that rule.');
  }

  var rejected = {};
  var rows = result.rows.filter(function (row) {
    for (var i = 0; i < active.length; i++) {
      var raw = row[active[i].idx];
      if (!active[i].accepted[norm_(raw)]) {
        var seen = String(raw == null ? '' : raw).trim() || '(blank)';
        var k = active[i].label + ' = ' + seen;
        rejected[k] = (rejected[k] || 0) + 1;
        return false;
      }
    }
    return true;
  });

  // Naming the rejected values makes a value mismatch obvious immediately.
  var reasons = Object.keys(rejected).sort(function (a, b) { return rejected[b] - rejected[a]; });
  if (reasons.length) {
    Logger.log('  Excluded ' + (result.rows.length - rows.length) + ' row(s): ' +
      reasons.slice(0, 6).map(function (r) { return r + ' ×' + rejected[r]; }).join(', ') +
      (reasons.length > 6 ? ', …' : ''));
  }
  return { cols: cols, rows: rows, filterError: null };
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
