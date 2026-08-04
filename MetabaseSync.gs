/* ═══════════════════════════════════════════════════════════════════════════
   Metabase → Google Sheets Sync
   meta.recykal.com  ·  Queries 5712 (Seller) & 5711 (Buyer)
   ─────────────────────────────────────────────────────────────────────────
   ONE-TIME SETUP  (do this before running):
     1. Find the line:  METABASE_PASS : 'YOUR_PASSWORD_HERE',
     2. Replace  YOUR_PASSWORD_HERE  with your actual Metabase password
     3. Save (Ctrl+S), then click ▶ Run → syncMetabaseToSheet

   The script will log progress in the Apps Script Execution log.
   ═══════════════════════════════════════════════════════════════════════════ */

// ── Configuration ────────────────────────────────────────────────────────────

var CFG = {
  METABASE_URL  : 'https://meta.recykal.com',
  METABASE_USER : 'vishwash.tiwari@recykal.com',
  METABASE_PASS : 'YOUR_PASSWORD_HERE',          // ← replace with your Metabase password
  SHEET_ID      : '1UMtuarqR9wFI74VM4JWC3GXSF9C9YpkkySeFJgq8rQc',
  QUERIES: [
    { id: 5712, tab: 'Seller' },
    { id: 5711, tab: 'Buyer'  },
  ],
  MAX_ROWS: 100000,   // safety cap — raise if you ever exceed this

  // Row-level filter applied after fetching — both conditions must match.
  // Column names are matched case-insensitively and fuzzy (spaces/underscores ignored).
  FILTER: {
    VERTICAL         : 'Open Marketplace',   // must equal Business Vertical / Vertical
    ONBOARDING_STATUS: 'Completed',          // must equal Onboarding Status
  },
};

// Required column order (must match header row in the sheet)
var COLUMN_ORDER = [
  'No.',
  'Seller Business Name',
  'Region',
  'Vertical',
  'Seller_Type',
  'State',
  'Seller_GSTIN',
  'Vintage_with_Recykal',
  'Finoscale_Rating',
  'Mail_ID',
  'POC_mail',
  'Mobile_No',
  'POC Names',
  'Date_of_Registration',
  'Entity Type',
  'Pending Document',
  'Debt Profile',
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


// ── Main entry point ─────────────────────────────────────────────────────────

function syncMetabaseToSheet() {
  var password = CFG.METABASE_PASS;
  if (!password || password === 'YOUR_PASSWORD_HERE') {
    throw new Error('Please replace YOUR_PASSWORD_HERE in CFG.METABASE_PASS with your actual Metabase password.');
  }

  Logger.log('Logging in to ' + CFG.METABASE_URL + ' as ' + CFG.METABASE_USER + ' …');
  var token = metabaseLogin_(CFG.METABASE_USER, password);
  Logger.log('Login OK. Session: ' + token.substring(0, 8) + '…');

  var ss = SpreadsheetApp.openById(CFG.SHEET_ID);

  CFG.QUERIES.forEach(function (q) {
    Logger.log('━━ Query ' + q.id + ' → "' + q.tab + '" tab ━━');
    var result = metabaseQueryFull_(token, q.id);
    Logger.log('   Received ' + result.rows.length + ' rows, ' + result.cols.length + ' columns');
    result = applyFilter_(result);
    Logger.log('   After filter (Vertical="' + CFG.FILTER.VERTICAL + '" + Onboarding Status="' + CFG.FILTER.ONBOARDING_STATUS + '"): ' + result.rows.length + ' rows');
    writeToSheet_(ss, q.tab, result);
    Logger.log('   ✓ "' + q.tab + '" updated');
  });

  Logger.log('━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━');
  Logger.log('Sync complete.');
}


// ── Metabase API helpers ──────────────────────────────────────────────────────

/**
 * Authenticate and return a session token.
 */
function metabaseLogin_(email, password) {
  var resp = UrlFetchApp.fetch(CFG.METABASE_URL + '/api/session', {
    method            : 'post',
    contentType       : 'application/json',
    payload           : JSON.stringify({ username: email, password: password }),
    muteHttpExceptions: true,
  });
  if (resp.getResponseCode() !== 200) {
    throw new Error(
      'Metabase login failed (HTTP ' + resp.getResponseCode() + ').\n' +
      'Check your email/password in Script properties.\n' +
      'Response: ' + resp.getContentText().substring(0, 400)
    );
  }
  return JSON.parse(resp.getContentText()).id;
}

/**
 * Fetch all rows for a saved question (card).
 * Strategy:
 *   1. GET /api/card/{id}          — retrieve the card's dataset_query
 *   2. POST /api/dataset            — run with a high max-results constraint
 * This bypasses the 2 000-row display cap that the UI applies.
 */
function metabaseQueryFull_(token, cardId) {
  // ── Step 1: get the card definition ───────────────────────────────────────
  var cardResp = UrlFetchApp.fetch(CFG.METABASE_URL + '/api/card/' + cardId, {
    method            : 'get',
    headers           : { 'X-Metabase-Session': token },
    muteHttpExceptions: true,
  });
  if (cardResp.getResponseCode() !== 200) {
    throw new Error(
      'Could not fetch card ' + cardId + ' (HTTP ' + cardResp.getResponseCode() + ').\n' +
      cardResp.getContentText().substring(0, 400)
    );
  }
  var card         = JSON.parse(cardResp.getContentText());
  var datasetQuery = card.dataset_query;
  if (!datasetQuery) throw new Error('Card ' + cardId + ' has no dataset_query — is it a saved question?');

  // ── Step 2: run the dataset query with an uncapped row limit ──────────────
  var queryPayload = JSON.stringify({
    database    : datasetQuery.database,
    type        : datasetQuery.type,
    query       : datasetQuery.query,
    native      : datasetQuery.native,
    parameters  : [],
    constraints : { 'max-results': CFG.MAX_ROWS },
  });

  var dataResp = UrlFetchApp.fetch(CFG.METABASE_URL + '/api/dataset', {
    method            : 'post',
    contentType       : 'application/json',
    headers           : { 'X-Metabase-Session': token },
    payload           : queryPayload,
    muteHttpExceptions: true,
  });
  var code = dataResp.getResponseCode();
  if (code !== 200 && code !== 202) {
    throw new Error(
      'Dataset query for card ' + cardId + ' failed (HTTP ' + code + ').\n' +
      dataResp.getContentText().substring(0, 400)
    );
  }

  var body = JSON.parse(dataResp.getContentText());
  var data = body.data || body;
  var cols = (data.cols || []).map(function (c) {
    return String(c.display_name || c.name || '').trim();
  });
  var rows = data.rows || [];

  if (data.rows_truncated != null && data.rows_truncated < rows.length) {
    Logger.log('   ⚠ rows_truncated=' + data.rows_truncated + ' — result may be incomplete. Raise CFG.MAX_ROWS.');
  }

  return { cols: cols, rows: rows };
}


// ── Row filter ────────────────────────────────────────────────────────────────

/**
 * Keep only rows where:
 *   Business Vertical / Vertical  =  CFG.FILTER.VERTICAL          ("Open Marketplace")
 *   Onboarding Status             =  CFG.FILTER.ONBOARDING_STATUS ("Completed")
 *
 * Column matching is case-insensitive and ignores spaces / underscores.
 * If either filter column is absent from the query output a warning is logged
 * and that condition is skipped (so you still get data rather than nothing).
 */
function applyFilter_(result) {
  var cols = result.cols;

  // Find column indices by normalised name
  function findCol(candidates) {
    for (var i = 0; i < cols.length; i++) {
      var n = norm_(cols[i]);
      for (var c = 0; c < candidates.length; c++) {
        if (n === norm_(candidates[c])) return i;
      }
    }
    return -1;
  }

  var verticalIdx  = findCol(['Vertical', 'Business Vertical', 'BusinessVertical']);
  var onboardingIdx = findCol(['Onboarding Status', 'OnboardingStatus', 'Onboarding_Status']);

  if (verticalIdx < 0) {
    Logger.log('   ⚠ "Vertical / Business Vertical" column not found in query output — vertical filter skipped.');
  }
  if (onboardingIdx < 0) {
    Logger.log('   ⚠ "Onboarding Status" column not found in query output — onboarding filter skipped.');
  }

  var wantVertical   = CFG.FILTER.VERTICAL.toLowerCase().trim();
  var wantOnboarding = CFG.FILTER.ONBOARDING_STATUS.toLowerCase().trim();

  var filtered = result.rows.filter(function (row) {
    var verticalOk   = verticalIdx  < 0 || String(row[verticalIdx]  || '').toLowerCase().trim() === wantVertical;
    var onboardingOk = onboardingIdx < 0 || String(row[onboardingIdx] || '').toLowerCase().trim() === wantOnboarding;
    return verticalOk && onboardingOk;
  });

  return { cols: cols, rows: filtered };
}


// ── Sheet writer ──────────────────────────────────────────────────────────────

/**
 * Map result columns → sheet columns and write rows starting at row 2.
 * Row 1 (headers) is never modified.
 * Columns in the sheet that have no matching Metabase column are left blank.
 */
function writeToSheet_(ss, tabName, result) {
  var sh = ss.getSheetByName(tabName);
  if (!sh) throw new Error('Tab "' + tabName + '" not found in the spreadsheet.');

  // ── Read existing headers ─────────────────────────────────────────────────
  var lastCol = sh.getLastColumn();
  var headers;
  if (lastCol > 0) {
    headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
      return String(h).trim();
    });
    // Drop trailing empty headers
    while (headers.length > 0 && headers[headers.length - 1] === '') headers.pop();
  }
  // Fall back to COLUMN_ORDER if row 1 is completely empty
  if (!headers || headers.length === 0) {
    Logger.log('   Row 1 of "' + tabName + '" is empty — writing COLUMN_ORDER as headers.');
    headers = COLUMN_ORDER.slice();
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  // ── Build a normalised lookup: query column name → column index in result ─
  var colLookup = {};
  result.cols.forEach(function (name, i) { colLookup[norm_(name)] = i; });

  // ── Map each sheet header to a result column (or -1 if unmatched) ─────────
  var mapping = headers.map(function (h) {
    var key = norm_(h);
    if (key in colLookup) return colLookup[key];
    // Try fuzzy: strip "seller" prefix (Buyer tab reuses same column order)
    var stripped = key.replace(/^seller/, '');
    for (var k in colLookup) {
      if (k === stripped || k.replace(/^seller/, '') === stripped) return colLookup[k];
    }
    return -1;
  });

  // ── Audit log ─────────────────────────────────────────────────────────────
  var unmatched = [];
  headers.forEach(function (h, c) {
    if (mapping[c] < 0) unmatched.push(h);
  });
  if (unmatched.length > 0) {
    Logger.log('   Columns not found in query output (will be blank): ' + unmatched.join(', '));
  }

  // ── Build output rows ─────────────────────────────────────────────────────
  var outRows = result.rows.map(function (srcRow) {
    return headers.map(function (_, c) {
      if (mapping[c] < 0) return '';
      var v = srcRow[mapping[c]];
      return v == null ? '' : v;
    });
  });

  // ── Clear old data (rows 2+), then write ──────────────────────────────────
  var lastRow = sh.getLastRow();
  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, headers.length).clearContent();
  }
  if (outRows.length > 0) {
    sh.getRange(2, 1, outRows.length, headers.length).setValues(outRows);
  }

  // ── Final validation log ──────────────────────────────────────────────────
  Logger.log('   Rows written : ' + outRows.length);
  Logger.log('   Columns matched: ' + mapping.filter(function (i) { return i >= 0; }).length +
             ' / ' + headers.length);
  Logger.log('   Sheet row count (incl. header): ' + (outRows.length + 1));
}


// ── Utility ───────────────────────────────────────────────────────────────────

/**
 * Normalise a column name for fuzzy matching:
 * lower-case, strip spaces / underscores / punctuation.
 */
function norm_(s) {
  return String(s).toLowerCase().replace(/[\s_\-\/(),\.]+/g, '').trim();
}
