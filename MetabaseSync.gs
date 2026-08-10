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
  SHEET_ID        : '1d57KGl00-pGWVjYKouyMu8jt0Y4UMEc2HaHWtMWPjeM',
  SYNC_EVERY_MINS : 1,
  QUERIES: [
    { id: 5712, tab: 'Sellers' },
    { id: 5711, tab: 'Buyers'  },
  ],
  FILTER: {
    VERTICAL         : 'Open Marketplace',
    ONBOARDING_STATUS: 'Completed',
  },
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

function syncMetabaseToSheet() {
  if (!CFG.METABASE_PASS || CFG.METABASE_PASS === 'YOUR_PASSWORD_HERE') {
    throw new Error('Set CFG.METABASE_PASS to your Metabase password and save the script.');
  }

  var token = getSessionToken_();
  var ss    = SpreadsheetApp.openById(CFG.SHEET_ID);

  CFG.QUERIES.forEach(function (q) {
    Logger.log('━━━ Query ' + q.id + ' → "' + q.tab + '" ━━━');

    var raw = fetchCardData_(token, q.id);
    Logger.log('Fetched: ' + raw.rows.length + ' rows, ' + raw.cols.length + ' cols');
    Logger.log('Metabase columns: ' + raw.cols.join(' | '));

    // Build GSTIN → earliest-date map from ALL rows (all verticals) BEFORE
    // applying the vertical filter. Vendors onboarded in another vertical first
    // get credit for their full Recykal tenure, not just their Open Marketplace date.
    var gstinDateMap = buildEarliestDateMap_(raw);

    var filtered = applyFilter_(raw);
    Logger.log('After filter: ' + filtered.rows.length + ' rows');

    if (filtered.rows.length === 0) {
      Logger.log('⚠ Zero rows after filter — check CFG.FILTER values match column names above');
    }

    writeToSheet_(ss, q.tab, filtered, gstinDateMap);
    formatSheet_(ss, q.tab, filtered.rows.length);
    Logger.log('✓ "' + q.tab + '" done');
  });

  Logger.log('══ Sync complete ══');
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

  Logger.log('  Filter — Vertical col idx: ' + vIdx + ', Onboarding col idx: ' + oIdx);
  if (vIdx < 0)  Logger.log('  ⚠ Vertical column not found — run debugColumns() to see exact names');
  if (oIdx < 0)  Logger.log('  ⚠ Onboarding Status column not found — run debugColumns() to see exact names');

  var wV = CFG.FILTER.VERTICAL.toLowerCase().trim();
  var wO = CFG.FILTER.ONBOARDING_STATUS.toLowerCase().trim();

  var filtered = result.rows.filter(function (row) {
    var vOk = (vIdx < 0) || String(row[vIdx] == null ? '' : row[vIdx]).toLowerCase().trim() === wV;
    var oOk = (oIdx < 0) || String(row[oIdx] == null ? '' : row[oIdx]).toLowerCase().trim() === wO;
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
