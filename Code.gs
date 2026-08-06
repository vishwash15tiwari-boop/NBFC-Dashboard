/**
 * ─────────────────────────────────────────────────────────────────────────────
 *  Recykal · NBFC Document Tracker (Web App backend)
 * ─────────────────────────────────────────────────────────────────────────────
 *  Single Google Sheet (SOURCE_FILE_ID) with three NBFC tracker tabs:
 *
 *    • "Billmart"    — entities submitting docs to Billmart NBFC
 *    • "Capital XB"  — entities submitting docs to Capital XB NBFC
 *    • "StrideOne"   — entities submitting docs to StrideOne NBFC
 *
 *  An optional "Seller Requirement" tab drives per-entity-type doc
 *  applicability (which docs are required for Proprietorship vs Partnership
 *  vs Private Limited). If that tab is absent, all docs are treated as
 *  required for every entity type — safe, conservative fallback.
 * ─────────────────────────────────────────────────────────────────────────────
 */

// Drive root folder where uploaded documents are organised.
// Structure: DRIVE_ROOT / <NBFC name> / <Entity name> / <file>
var DRIVE_ROOT_ID = '1i5melXCocWrV9rR-3gM75wwSwWy7Dqit';

// Marketplace sheet — source of truth for total active sellers & buyers.
// Only rows where the Vertical column = MB_VERTICAL_FILTER are counted.
var MB_SHEET_ID        = '1d57KGl00-pGWVjYKouyMu8jt0Y4UMEc2HaHWtMWPjeM';
var MB_VERTICAL_FILTER = 'Open Marketplace';

var CONFIG = {
  // Native Google Sheet ID (already confirmed native; no xlsx conversion needed).
  SOURCE_FILE_ID: '1RoHWbZyHhNKlweWXD4AMSZfB5ONdktPcVayOkpPgjpo',

  // One entry per NBFC tracker tab.
  // maxCol caps how far right the tab is read so scratch columns to the right
  // never pollute KPIs, cards, the matrix, or form saves.
  // docStartCol / docEndCol (1-indexed) restrict which columns are treated as
  // document-status columns — L=12 through AA=27.
  NBFC_TABS: [
    { id: 'billmart',  name: 'Billmart',   maxCol: 40, docStartCol: 12, docEndCol: 27 },
    { id: 'capitalxb', name: 'Capital XB', maxCol: 40, docStartCol: 12, docEndCol: 27 },
    { id: 'strideone', name: 'StrideOne',  maxCol: 40, docStartCol: 12, docEndCol: 27 },
  ],

  // Funnel-stage columns (1-indexed): AG=33, AH=34, AI=35.
  // These are fixed positions in all three NBFC tabs.
  FUNNEL_COLS: { eligibility: 33, qualified: 34, creditLimit: 35 },

  // Optional: entity-type → document applicability matrix tab.
  REQUIREMENT: { name: 'Seller Requirement', index: -1, signature: ['documents', 'proprietor'] },

  // Seller follow-up remarks tab. One row per seller (keyed by Seller_GSTIN)
  // carrying an ETA (expected date) and a free-text Remarks note. Joined onto
  // every seller row by GSTIN so remarks surface per seller in the dashboard.
  REMARKS: { name: 'Plastic-Remarks', index: -1, signature: ['remark', 'gst'] },

  APP_TITLE: 'Recykal · NBFC Document Tracker',
  PROP_BACKEND_ID: 'BACKEND_SHEET_ID'
};

/* ── Columns to hide from the document matrix and detail modal (all tabs) ──
   Matched case-insensitively as substrings of the Google Sheet column header.
   Each keyword is unique to the non-standard columns and absent from all
   standard docs (GST Certificate, PAN Card, Aadhaar Card, etc.). */
var HIDDEN_DOC_KEYS_ = [
  'electricity',       // Electricity Bill / Rental Agreement
  'rental agreement',  // Electricity Bill / Rental Agreement (alt)
  'credit approv',     // Credit Approved / Credit Approval
  'director',          // Aadhar (Owner/Director/Partner) + Owner/Director/Partner PAN
  'billmart qualif',   // Billmart Qualified
  'shareholding',      // Shareholding Details
  ', coi',             // MOA, AOA , COI  (comma distinguishes it from MOA & AOA)
  'moa, aoa',          // MOA, AOA variant
  'moa,aoa'            // MOA,AOA no-space variant
];
function isHiddenDoc_(key) {
  var k = String(key || '').toLowerCase().trim();
  return HIDDEN_DOC_KEYS_.some(function (h) { return k.indexOf(h) !== -1; });
}

/* ─────────────────────────── Web-app entry ─────────────────────────── */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(CONFIG.APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/* ─────────────────────────── Spreadsheet access ────────────────────────── */

function getSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();
  var savedId = props.getProperty(CONFIG.PROP_BACKEND_ID);
  if (savedId) {
    try { return SpreadsheetApp.openById(savedId); } catch (e) { /* stale */ }
  }
  try {
    var direct = SpreadsheetApp.openById(CONFIG.SOURCE_FILE_ID);
    props.setProperty(CONFIG.PROP_BACKEND_ID, direct.getId());
    return direct;
  } catch (e) { /* not native — convert once */ }
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    savedId = props.getProperty(CONFIG.PROP_BACKEND_ID);
    if (savedId) {
      try { return SpreadsheetApp.openById(savedId); } catch (e2) { /* fall through */ }
    }
    var newId = convertExcelToNativeSheet_(CONFIG.SOURCE_FILE_ID);
    props.setProperty(CONFIG.PROP_BACKEND_ID, newId);
    return SpreadsheetApp.openById(newId);
  } finally {
    lock.releaseLock();
  }
}

function convertExcelToNativeSheet_(fileId) {
  var file = DriveApp.getFileById(fileId);
  var blob = file.getBlob();
  var meta = {
    name: file.getName().replace(/\.xlsx?$/i, '') + ' (Live)',
    mimeType: 'application/vnd.google-apps.spreadsheet'
  };
  var parents = file.getParents();
  if (parents.hasNext()) meta.parents = [parents.next().getId()];
  var boundary = 'rkboundary' + new Date().getTime();
  var head = '--' + boundary + '\r\n' +
             'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
             JSON.stringify(meta) + '\r\n' +
             '--' + boundary + '\r\n' +
             'Content-Type: ' + blob.getContentType() + '\r\n\r\n';
  var tail = '\r\n--' + boundary + '--';
  var payload = Utilities.newBlob(head).getBytes()
    .concat(blob.getBytes())
    .concat(Utilities.newBlob(tail).getBytes());
  var res = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true',
    { method: 'post', contentType: 'multipart/related; boundary=' + boundary,
      payload: payload, headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true }
  );
  if (res.getResponseCode() >= 300)
    throw new Error('Could not convert the Excel workbook. Drive API said: ' + res.getContentText());
  return JSON.parse(res.getContentText()).id;
}

/* ─────────────────────────── Tab discovery ─────────────────────────── */

function normKey_(s) {
  return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]+/g, '');
}

function findSheet_(ss, spec) {
  var sheets = ss.getSheets();
  var want = normKey_(spec.name);
  var i, sh;
  for (i = 0; i < sheets.length; i++) {
    if (normKey_(sheets[i].getName()) === want) return sheets[i];
  }
  if (spec.signature) {
    for (i = 0; i < sheets.length; i++) {
      sh = sheets[i];
      if (sh.getLastRow() < 1 || sh.getLastColumn() < 1) continue;
      var header = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0]
        .map(function (h) { return String(h).toLowerCase(); }).join('|');
      var hit = spec.signature.every(function (tok) { return header.indexOf(tok) !== -1; });
      if (hit) return sh;
    }
  }
  if (spec.index >= 0 && spec.index < sheets.length) return sheets[spec.index];
  throw new Error('Could not locate the "' + spec.name + '" tab in the spreadsheet.');
}

/* ─────────────────────────── Status vocab ─────────────────────────── */

function parseStatus_(raw) {
  var s = String(raw == null ? '' : raw).trim();
  var l = s.toLowerCase();
  if (!s || s === '-') return { status: 'pending', note: '' };
  if (l === 'na' || l === 'n/a' || l === 'not applicable') return { status: 'na', note: '' };
  if (l.indexOf('pending') !== -1) return { status: 'pending', note: l === 'pending' ? '' : s };
  if (l === 'received' || l === 'yes' || l === 'done' || l === 'submitted' ||
      l.indexOf('receiv') !== -1 || l.indexOf('provided') !== -1)
    return { status: 'received', note: (l === 'received' || l === 'yes') ? '' : s };
  return { status: 'pending', note: s };
}

/* ──────────────────── Hardcoded entity-type doc applicability ──────────── */
//
// Each entry: [ normalisedDocFragment, [entityClasses where doc is N/A] ]
//
// Applicability by category (confirmed requirements):
//   Proprietorship  — N/A: Partnership Deed, MOA/AOA/COI, Shareholding Details
//   Partnership     — N/A: MOA/AOA/COI
//   Private Limited — N/A: Partnership Deed
//
var ENTITY_DOC_RULES = [
  ['partnershipdeed',         ['proprietorship', 'privatelimited']],
  ['moaaoacoi',               ['proprietorship', 'partnership']],
  ['memorandumofassociation', ['proprietorship', 'partnership']],
  ['articleofassociation',    ['proprietorship', 'partnership']],
  ['certificateofincorporat', ['proprietorship', 'partnership']],
  ['shareholdingdetail',      ['proprietorship']],
  ['shareholdingpattern',     ['proprietorship']],
];

// Documents that are ALWAYS required for every entity type — never marked N/A and
// never overridden by the optional "Seller Requirement" matrix tab.
var ALWAYS_REQUIRED_DOC_FRAGMENTS = ['debtprofile'];

/**
 * Maps a raw entity-type string to one of three canonical classes:
 *   'proprietorship' | 'partnership' | 'privatelimited'
 * Returns null when the type is blank or unrecognised.
 * Private Limited is tested first because its normalised form contains
 * "limited" which would also partially match LLP / "limitedliabilitypartnership".
 */
function classifyEntityType_(entityType) {
  var n = normKey_(entityType);
  if (!n) return null;
  if (n.indexOf('privat') !== -1 || (n.indexOf('pvt') !== -1 && (n.indexOf('ltd') !== -1 || n.indexOf('lim') !== -1))) return 'privatelimited';
  if (n.indexOf('partner') !== -1 || n.indexOf('llp') !== -1) return 'partnership';
  if (n.indexOf('proprietor') !== -1 || n.indexOf('propri') !== -1) return 'proprietorship';
  return null;
}

/**
 * Returns false when ENTITY_DOC_RULES says this document is not applicable for
 * the given entity class.  Returns true (applicable) when the entity class is
 * unknown or no rule matches the document header.
 */
function isDocApplicableByRules_(docHeader, entityClass) {
  var docNorm = normKey_(docHeader);
  // Always-required docs bypass all entity-type exclusion rules.
  for (var j = 0; j < ALWAYS_REQUIRED_DOC_FRAGMENTS.length; j++) {
    if (docNorm.indexOf(ALWAYS_REQUIRED_DOC_FRAGMENTS[j]) !== -1) return true;
  }
  if (!entityClass) return true;
  for (var i = 0; i < ENTITY_DOC_RULES.length; i++) {
    var fragment  = ENTITY_DOC_RULES[i][0];
    var naClasses = ENTITY_DOC_RULES[i][1];
    if (docNorm.indexOf(fragment) !== -1) {
      if (naClasses.indexOf(entityClass) !== -1) return false;
    }
  }
  return true;
}

/* ──────────────────────── Requirement matrix (optional tab) ───────────── */

function readRequirementMatrix_(ss) {
  try {
    var sh = findSheet_(ss, CONFIG.REQUIREMENT);
    var values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues();
    var headerRow = -1, docCol = -1, r, c;
    for (r = 0; r < values.length && headerRow === -1; r++) {
      for (c = 0; c < values[r].length; c++) {
        if (normKey_(values[r][c]) === 'documents') { headerRow = r; docCol = c; break; }
      }
    }
    if (headerRow === -1) return { entityColumns: [], rows: [] };
    var entityCols = [];
    for (c = docCol + 1; c < values[headerRow].length; c++) {
      var label = String(values[headerRow][c]).trim();
      if (label) entityCols.push({ col: c, label: label });
    }
    var rows = [];
    for (r = headerRow + 1; r < values.length; r++) {
      var docName = String(values[r][docCol]).trim();
      if (!docName) continue;
      var requiredBy = {};
      entityCols.forEach(function (ec) {
        var v = String(values[r][ec.col]).trim();
        requiredBy[ec.label] = !!v && v !== '-' && v.toLowerCase() !== 'no' && v.toLowerCase() !== 'na';
      });
      rows.push({ name: docName, requiredBy: requiredBy });
    }
    return { entityColumns: entityCols.map(function (ec) { return ec.label; }), rows: rows };
  } catch (e) {
    return { entityColumns: [], rows: [] };
  }
}

function matchRequirementRow_(matrix, trackerDocHeader) {
  var target = normKey_(trackerDocHeader);
  var best = null, bestLen = 0;
  matrix.rows.forEach(function (row) {
    var n = normKey_(row.name);
    if (!n) return;
    if (n === target || target.indexOf(n) !== -1 || n.indexOf(target) !== -1) {
      if (n.length > bestLen) { best = row; bestLen = n.length; }
    }
  });
  return best;
}

function matchEntityColumn_(matrix, entityType) {
  var e = normKey_(entityType);
  if (!e) return null;
  var best = null, bestLen = 0, exactFound = false;
  matrix.entityColumns.forEach(function (label) {
    var n = normKey_(label);
    if (n === e) {
      if (!exactFound || n.length > bestLen) { best = label; bestLen = n.length; exactFound = true; }
    } else if (!exactFound && (e.indexOf(n) !== -1 || n.indexOf(e) !== -1) && n.length > bestLen) {
      best = label; bestLen = n.length;
    }
  });
  return best;
}

/* ──────────────────── Seller remarks (Plastic-Remarks tab) ─────────────── */

// Normalise a GSTIN for matching: strip whitespace, upper-case.
function normGst_(s) {
  return String(s == null ? '' : s).replace(/\s+/g, '').toUpperCase();
}

/**
 * Reads the "Plastic-Remarks" tab and returns a map keyed by normalised GSTIN:
 *   { <normGST>: { remarks: <string>, eta: <string> } }
 *
 * The tab is a manually-maintained follow-up tracker — one row per seller
 * (identified by Seller_GSTIN) carrying an ETA (expected date) and a free-text
 * Remarks note. Rows without a GSTIN, or with neither a remark nor an ETA, are
 * skipped. When the same GSTIN appears more than once, the last populated row
 * wins (latest entry in the sheet). On any error (tab missing / unreadable) an
 * empty map is returned so the dashboard still loads.
 */
function getRemarksMap_(ss) {
  var map = {};
  try {
    var sh = findSheet_(ss, CONFIG.REMARKS);
    var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return map;
    var values = sh.getRange(1, 1, lastRow, lastCol).getDisplayValues();

    // Locate the header row (scan the first 3 rows) and the columns we need.
    var headerRow = -1, gstIdx = -1, etaIdx = -1, remarksIdx = -1;
    var maxScan = Math.min(3, values.length);
    for (var r = 0; r < maxScan; r++) {
      var gi = -1, ei = -1, ri = -1;
      for (var c = 0; c < values[r].length; c++) {
        var nk = normKey_(values[r][c]);
        if (gi === -1 && nk.indexOf('gst')    !== -1) gi = c;
        if (ei === -1 && nk === 'eta')                ei = c;
        if (ri === -1 && nk.indexOf('remark') !== -1) ri = c;
      }
      if (ri !== -1 && gi !== -1) { headerRow = r; gstIdx = gi; etaIdx = ei; remarksIdx = ri; break; }
    }
    if (headerRow === -1) return map;

    for (var d = headerRow + 1; d < values.length; d++) {
      var gst = normGst_(values[d][gstIdx]);
      if (!gst) continue;
      var remark = remarksIdx >= 0 ? String(values[d][remarksIdx] == null ? '' : values[d][remarksIdx]).trim() : '';
      var eta    = etaIdx    >= 0 ? String(values[d][etaIdx]    == null ? '' : values[d][etaIdx]).trim()    : '';
      if (!remark && !eta) continue;               // nothing to surface for this seller
      var prev = map[gst];                         // last populated row wins; keep prior non-empty fields
      map[gst] = {
        remarks: remark || (prev ? prev.remarks : ''),
        eta:     eta    || (prev ? prev.eta     : '')
      };
    }
  } catch (e) { /* tab missing or unreadable — surface no remarks rather than fail */ }
  return map;
}

/* ─────────────────────────── Tracker layout ─────────────────────────── */

function readTrackerLayout_(sh, maxCol, docStartCol, docEndCol) {
  var lastCol = sh.getLastColumn();
  if (maxCol && maxCol > 0 && maxCol < lastCol) lastCol = maxCol;
  // Scan up to 3 rows to find the actual header row (handles sheets where row 1
  // is a title or blank and column headers start at row 2 or 3).
  var headerRowNum = 0, headers = [], pendingIdx = -1, i;
  var maxSearch = Math.min(3, sh.getLastRow()), r, nk;
  for (r = 1; r <= maxSearch && !headerRowNum; r++) {
    var candidate = sh.getRange(r, 1, 1, lastCol).getDisplayValues()[0]
      .map(function (h) { return String(h).trim(); });
    for (i = 0; i < candidate.length; i++) {
      nk = normKey_(candidate[i]);
      if (nk === 'pendingdocument' || nk === 'pendingdocuments') {
        headerRowNum = r; headers = candidate; pendingIdx = i; break;
      }
    }
  }
  if (!headerRowNum)
    throw new Error('The "' + sh.getName() + '" tab has no "Pending Document" column in the first 3 rows.');
  var serialIdx = 0;
  var metaIdx = [];
  for (i = 0; i < pendingIdx; i++) {
    if (i === serialIdx) continue;
    if (headers[i]) metaIdx.push(i);
  }
  // Only include doc columns within the L:AA window (1-indexed docStartCol to docEndCol).
  // When no window is configured, all non-empty columns after pendingIdx qualify.
  var docIdx = [];
  for (i = pendingIdx + 1; i < headers.length; i++) {
    if (!headers[i]) continue;
    var col1 = i + 1; // 1-indexed column number
    if (docStartCol && docEndCol) {
      if (col1 >= docStartCol && col1 <= docEndCol) docIdx.push(i);
    } else {
      docIdx.push(i);
    }
  }
  // Columns after docEndCol are display-only seller attributes (e.g. Vintage, Eligibility).
  var extraMetaIdx = [];
  if (docEndCol) {
    for (i = 0; i < headers.length; i++) {
      if (headers[i] && (i + 1) > docEndCol) extraMetaIdx.push(i);
    }
  }
  return { headers: headers, colCount: lastCol, headerRowNum: headerRowNum, serialIdx: serialIdx, metaIdx: metaIdx, pendingIdx: pendingIdx, docIdx: docIdx, extraMetaIdx: extraMetaIdx };
}

function metaFieldType_(header) {
  var n = normKey_(header);
  if (n.indexOf('gst') !== -1 && n.indexOf('return') === -1 && n.indexOf('certificate') === -1) return 'gst';
  if (n.indexOf('date') !== -1) return 'date';
  if (n.indexOf('entitytype') !== -1 || n.indexOf('businesstype') !== -1 ||
      n.indexOf('typeofentity') !== -1 || n.indexOf('typeofbusiness') !== -1 ||
      n.indexOf('typeoforganis') !== -1 || n.indexOf('organisationtype') !== -1 ||
      n.indexOf('organizationtype') !== -1 || n.indexOf('companytype') !== -1 ||
      n.indexOf('entitycategory') !== -1 || n.indexOf('businesscategory') !== -1 ||
      n === 'entity' || n === 'entitycategory') return 'entity';
  if (n.indexOf('turnover') !== -1 || n === 'annualrevenue' || n === 'revenue') return 'turnover';
  return 'text';
}

/* ────────────────────────── Per-NBFC data reader ──────────────────────── */

/**
 * Reads one NBFC tracker tab and returns all entity rows with parsed doc statuses.
 * On any error, returns an error object so getInitialData() can still succeed.
 */
function getNbfcData_(ss, tabCfg, matrix, remarksMap) {
  try {
    var sh = findSheet_(ss, tabCfg);
    var layout = readTrackerLayout_(sh, tabCfg.maxCol, tabCfg.docStartCol, tabCfg.docEndCol);
    remarksMap = remarksMap || {};

    var nameHeaderKey = null, entityHeaderKey = null, gstHeaderKey = null;
    layout.metaIdx.forEach(function (idx) {
      var h = layout.headers[idx];
      if (!nameHeaderKey && normKey_(h).indexOf('name') !== -1) nameHeaderKey = h;
      if (!entityHeaderKey && metaFieldType_(h) === 'entity') entityHeaderKey = h;
      if (!gstHeaderKey && metaFieldType_(h) === 'gst') gstHeaderKey = h;
    });

    // Build doc list — match each column against the requirement matrix
    var docs = layout.docIdx.filter(function (idx) {
      return !isHiddenDoc_(layout.headers[idx]);
    }).map(function (idx) {
      var header = layout.headers[idx];
      var req = matchRequirementRow_(matrix, header);
      var requiredBy = {};
      matrix.entityColumns.forEach(function (label) {
        requiredBy[label] = req ? !!req.requiredBy[label] : true;
      });
      return { key: header, requiredBy: requiredBy };
    });

    var lastRow = sh.getLastRow();
    var entities = [];
    var optionValues = {};

    if (lastRow > layout.headerRowNum) {
      var dataRange = sh.getRange(layout.headerRowNum + 1, 1, lastRow - layout.headerRowNum, layout.colCount);
      var values = dataRange.getDisplayValues();
      var allNotes = dataRange.getNotes();

      values.forEach(function (row, i) {
        // Validate the serial (No.) column.  Only count rows that have a
        // positive integer serial — blank rows, header echo-rows ("No:", "S.No"),
        // and rows with serial ≤ 0 (data errors / sentinel values) are all skipped.
        var serialRaw = String(row[layout.serialIdx] == null ? '' : row[layout.serialIdx]).trim();
        if (!serialRaw) return;                                                    // blank serial → skip
        var serialParsed = parseInt(serialRaw, 10);
        if (isNaN(serialParsed) && /[a-zA-Z]/.test(serialRaw)) return;            // "No:", "S.No" → skip
        if (!isNaN(serialParsed) && serialParsed < 1) return;                     // 0 or negative → skip

        var meta = {};
        layout.metaIdx.forEach(function (idx) { meta[layout.headers[idx]] = String(row[idx]).trim(); });
        var hasIdentity = layout.metaIdx.some(function (idx) { return String(row[idx]).trim() !== ''; });
        if (!hasIdentity) return;
        // Validate the seller/buyer name column:
        //   • must be present and non-empty
        //   • must not be purely numeric (e.g. "1") — those are data artifacts, not real entity names
        var sellerName = nameHeaderKey ? String(meta[nameHeaderKey] || '').trim() : '';
        if (nameHeaderKey && !sellerName) return;
        if (nameHeaderKey && /^\d+(\.\d+)?$/.test(sellerName)) return;

        layout.metaIdx.forEach(function (idx) {
          var h = layout.headers[idx], v = String(row[idx]).trim();
          if (!v) return;
          (optionValues[h] = optionValues[h] || {})[v] = true;
        });

        var serial = isNaN(serialParsed) ? '' : serialParsed;
        var entityType = entityHeaderKey ? (meta[entityHeaderKey] || '') : '';
        var entityCol   = matchEntityColumn_(matrix, entityType);
        var entityClass = classifyEntityType_(entityType);
        var rowNotes = allNotes[i] || [];
        var docStates = {};
        var received = 0, pending = 0, na = 0;
        var extraMeta = {};
        layout.extraMetaIdx.forEach(function (idx) {
          extraMeta[layout.headers[idx]] = String(row[idx] == null ? '' : row[idx]).trim();
        });

        layout.docIdx.forEach(function (idx) {
          var h = layout.headers[idx];
          var req = matchRequirementRow_(matrix, h);
          // Hardcoded entity-type rules take first priority.
          // Always-required docs (e.g. Debt Profile) are never overridden by the matrix.
          // The optional sheet matrix can further refine when both entity class
          // and a matching requirement row are present.
          var docNormH = normKey_(h);
          var alwaysRequired = ALWAYS_REQUIRED_DOC_FRAGMENTS.some(function(f){ return docNormH.indexOf(f) !== -1; });
          var applicable;
          if (!isDocApplicableByRules_(h, entityClass)) {
            applicable = false;
          } else if (!alwaysRequired && req && entityCol) {
            applicable = !!req.requiredBy[entityCol];
          } else {
            applicable = true;
          }
          var cellNote = String(rowNotes[idx] || '').trim();
          var driveUrl = /^https?:\/\/\S+$/.test(cellNote) ? cellNote : '';
          var parsed = parseStatus_(row[idx]);
          if (!applicable && parsed.status !== 'received') {
            docStates[h] = { status: 'na', note: '', raw: String(row[idx]).trim(), driveUrl: '' };
            na++;
          } else {
            docStates[h] = { raw: String(row[idx]).trim(), status: parsed.status, note: parsed.note, driveUrl: driveUrl };
            if (parsed.status === 'received') received++;
            else if (parsed.status === 'na') na++;
            else pending++;
          }
        });

        var applicable = received + pending;
        // Join the follow-up remark/ETA for this seller by GSTIN.
        var rk = gstHeaderKey ? remarksMap[normGst_(meta[gstHeaderKey])] : null;
        entities.push({
          row: i + layout.headerRowNum + 1,
          serial: isNaN(serial) ? '' : serial,
          meta: meta,
          extraMeta: extraMeta,
          docs: docStates,
          received: received,
          pending: pending,
          na: na,
          applicable: applicable,
          completion: applicable ? Math.round((received / applicable) * 1000) / 10 : 0,
          remarks: rk ? rk.remarks : '',
          eta:     rk ? rk.eta     : ''
        });
      });
    }

    // Merge entity type options: actual sheet values first, then matrix columns
    var entityOptions = {};
    entities.forEach(function (e) {
      if (entityHeaderKey && e.meta[entityHeaderKey]) entityOptions[e.meta[entityHeaderKey]] = true;
    });
    matrix.entityColumns.forEach(function (label) {
      var n = normKey_(label);
      var covered = Object.keys(entityOptions).some(function (existing) {
        var e = normKey_(existing);
        return e.indexOf(n) !== -1 || n.indexOf(e) !== -1;
      });
      if (!covered) entityOptions[label] = true;
    });

    // Read the three funnel-stage column headers by their fixed 1-indexed positions.
    var fc = CONFIG.FUNNEL_COLS;
    var eligibilityHeader = fc.eligibility  ? (layout.headers[fc.eligibility  - 1] || null) : null;
    var qualifiedHeader   = fc.qualified    ? (layout.headers[fc.qualified    - 1] || null) : null;
    var creditLimitHeader = fc.creditLimit  ? (layout.headers[fc.creditLimit  - 1] || null) : null;

    return {
      ok: true,
      id: tabCfg.id,
      name: tabCfg.name,
      entities: entities,
      docs: docs,
      entityColumns: matrix.entityColumns,
      entityTypeOptions: Object.keys(entityOptions).sort(),
      nameHeader: nameHeaderKey,
      entityHeader: entityHeaderKey,
      metaFields: layout.metaIdx.map(function (idx) {
        var h = layout.headers[idx];
        return { key: h, type: metaFieldType_(h), options: Object.keys(optionValues[h] || {}).sort() };
      }),
      extraMetaFields: layout.extraMetaIdx.map(function (idx) { return layout.headers[idx]; }),
      pendingHeader: layout.headers[layout.pendingIdx],
      eligibilityHeader: eligibilityHeader,
      qualifiedHeader:   qualifiedHeader,
      creditLimitHeader: creditLimitHeader
    };
  } catch (e) {
    return {
      ok: false, error: String(e.message),
      id: tabCfg.id, name: tabCfg.name,
      entities: [], docs: [], entityColumns: [], entityTypeOptions: [],
      nameHeader: null, entityHeader: null, metaFields: [],
      extraMetaFields: [], pendingHeader: null,
      eligibilityHeader: null, qualifiedHeader: null, creditLimitHeader: null
    };
  }
}

/* ──────────── Marketplace counts (Open Marketplace · Completed) ─── */

var MB_STATUS_FILTER = 'Completed';

/**
 * Counts rows in a marketplace tab where:
 *   Vertical column = MB_VERTICAL_FILTER ("Open Marketplace")
 *   AND Onboarding column (header "Onboarding", fallback col F) = MB_STATUS_FILTER ("Completed")
 *
 * Returns _debug so the browser console shows all column headers + every
 * distinct value in the matched onboarding column — use this to verify filters.
 */

/* ─────────────────── StrideOne buyer tracker (dedicated reader) ────────── */

/**
 * Reads the StrideOne tab, which tracks BUYERS with a layout unlike the seller
 * tabs — a "Priority" column (P0 = top buyer) instead of a serial, no "Pending
 * Document" column, buyer-named identity columns, and its own funnel
 * (Eligibility / Qualified / Credit Limit) + document set (MOA … Agings).
 *
 * Returns the SAME shape as getNbfcData_() so the frontend renders StrideOne
 * buyers with full detail (overview table, detail view, KPI drawers), plus:
 *   phase0        — count of P-0 priority buyers
 *   phase0Buyers  — P-0 buyer objects for the "Phase 0" drawer list
 *   _isBuyer      — flags the buyer-oriented tab to the frontend
 * On any error a valid empty object (ok:false) is returned so the dashboard
 * still loads.
 */
function getStrideOneData_(ss, remarksMap) {
  remarksMap = remarksMap || {};
  var tabCfg = CONFIG.NBFC_TABS[2];
  try {
    var sh = findSheet_(ss, tabCfg);
    var lastRow = sh.getLastRow(), lastColAll = sh.getLastColumn();
    if (lastRow < 2 || lastColAll < 1) throw new Error('The StrideOne tab is empty.');
    var lastCol = (tabCfg.maxCol && tabCfg.maxCol < lastColAll) ? tabCfg.maxCol : lastColAll;
    var range    = sh.getRange(1, 1, lastRow, lastCol);
    var values   = range.getDisplayValues();
    var allNotes = range.getNotes();

    // Find the buyer-table header row (scan the first 5 rows): needs a "Priority"
    // column and at least one "Buyer" column.
    var headerRowNum = -1, headers = null;
    for (var r = 0; r < Math.min(5, values.length); r++) {
      var keys = values[r].map(function (h) { return normKey_(h); });
      if (keys.indexOf('priority') !== -1 && keys.some(function (n) { return n.indexOf('buyer') !== -1; })) {
        headerRowNum = r; headers = values[r].map(function (h) { return String(h).trim(); }); break;
      }
    }
    if (headerRowNum === -1)
      throw new Error('The "StrideOne" tab has no buyer table (Priority / Buyer columns) in the first 5 rows.');

    function findCol(pred) { for (var i = 0; i < headers.length; i++) { if (headers[i] && pred(normKey_(headers[i]))) return i; } return -1; }
    var priorityIdx = findCol(function (n) { return n === 'priority'; });
    var nameIdx     = findCol(function (n) { return n.indexOf('businessname') !== -1; });
    var gstIdx      = findCol(function (n) { return n.indexOf('gst') !== -1; });
    var entityIdx   = findCol(function (n) { return n.indexOf('registrationtype') !== -1; });
    var catIdx      = findCol(function (n) { return n.indexOf('category') !== -1; });
    var stateIdx    = findCol(function (n) { return n === 'state'; });
    var contactIdx  = findCol(function (n) { return n.indexOf('contact') !== -1; });
    var vintageIdx  = findCol(function (n) { return n.indexOf('vintage') !== -1; });
    var eligIdx     = findCol(function (n) { return n.indexOf('eligib') !== -1; });
    var qualIdx     = findCol(function (n) { return n.indexOf('qualif') !== -1; });
    var clIdx       = findCol(function (n) { return n.indexOf('creditlimit') !== -1; });
    var emailStIdx  = findCol(function (n) { return n.indexOf('emailstatus') !== -1; });
    if (nameIdx === -1) throw new Error('The "StrideOne" tab has no Buyer Business Name column.');

    // Column roles:
    //   meta      — identity/attribute columns before the funnel block
    //   extraMeta — the funnel block: Eligibility, Qualified, Credit Limit, Email Status
    //   docs      — document columns after the funnel block (MOA … Agings)
    var funnelStart = eligIdx !== -1 ? eligIdx : headers.length;
    var lastFunnel  = Math.max(clIdx, emailStIdx, qualIdx, eligIdx);
    var docStart    = lastFunnel !== -1 ? lastFunnel + 1 : funnelStart;
    var metaIdx = [], extraMetaIdx = [], docIdx = [];
    for (var c = 0; c < headers.length; c++) {
      if (!headers[c]) continue;
      if (c < funnelStart)   metaIdx.push(c);
      else if (c < docStart) extraMetaIdx.push(c);
      else                   docIdx.push(c);
    }

    var docs = docIdx.filter(function (idx) {
      return !isHiddenDoc_(headers[idx]);
    }).map(function (idx) { return { key: headers[idx], requiredBy: {} }; });

    var entities = [], phase0 = 0, phase0Buyers = [], optionValues = {}, entityOptions = {};
    for (var d = headerRowNum + 1; d < values.length; d++) {
      var row = values[d];
      var name = String(row[nameIdx] == null ? '' : row[nameIdx]).trim();
      if (!name || /^\d+(\.\d+)?$/.test(name)) continue;   // skip blank / numeric-artifact rows

      var meta = {};
      metaIdx.forEach(function (idx) { meta[headers[idx]] = String(row[idx] == null ? '' : row[idx]).trim(); });
      var extraMeta = {};
      extraMetaIdx.forEach(function (idx) { extraMeta[headers[idx]] = String(row[idx] == null ? '' : row[idx]).trim(); });
      metaIdx.forEach(function (idx) { var h = headers[idx], v = meta[h]; if (v) (optionValues[h] = optionValues[h] || {})[v] = true; });

      var rowNotes = allNotes[d] || [];
      var docStates = {}, received = 0, pending = 0, na = 0;
      docIdx.forEach(function (idx) {
        var h = headers[idx];
        var cellNote = String(rowNotes[idx] || '').trim();
        var driveUrl = /^https?:\/\/\S+$/.test(cellNote) ? cellNote : '';
        var parsed = parseStatus_(row[idx]);
        docStates[h] = { raw: String(row[idx] == null ? '' : row[idx]).trim(), status: parsed.status, note: parsed.note, driveUrl: driveUrl };
        if (parsed.status === 'received') received++;
        else if (parsed.status === 'na') na++;
        else pending++;
      });
      var applicable = received + pending;

      var priority = priorityIdx !== -1 ? String(row[priorityIdx] || '').trim() : '';
      var isP0 = normKey_(priority).indexOf('p0') === 0;
      if (isP0) phase0++;
      if (entityIdx !== -1 && meta[headers[entityIdx]]) entityOptions[meta[headers[entityIdx]]] = true;

      var rk = gstIdx !== -1 ? remarksMap[normGst_(row[gstIdx])] : null;
      entities.push({
        row: d + 1,
        serial: entities.length + 1,
        priority: priority,
        meta: meta,
        extraMeta: extraMeta,
        docs: docStates,
        received: received, pending: pending, na: na, applicable: applicable,
        completion: applicable ? Math.round((received / applicable) * 1000) / 10 : 0,
        remarks: rk ? rk.remarks : '',
        eta:     rk ? rk.eta     : ''
      });

      if (isP0) phase0Buyers.push({
        row:      d + 1,
        name:     name,
        gst:      gstIdx     >= 0 ? String(row[gstIdx]     || '').trim() : '',
        category: catIdx     >= 0 ? String(row[catIdx]     || '').trim() : '',
        state:    stateIdx   >= 0 ? String(row[stateIdx]   || '').trim() : '',
        contact:  contactIdx >= 0 ? String(row[contactIdx] || '').trim() : '',
        type:     entityIdx  >= 0 ? String(row[entityIdx]  || '').trim() : '',
        vintage:  vintageIdx >= 0 ? String(row[vintageIdx] || '').trim() : '',
        priority: priority
      });
    }

    return {
      ok: true, id: 'strideone', name: 'StrideOne',
      entities: entities, docs: docs,
      entityColumns: [],
      entityTypeOptions: Object.keys(entityOptions).sort(),
      nameHeader: headers[nameIdx],
      entityHeader: entityIdx !== -1 ? headers[entityIdx] : null,
      metaFields: metaIdx.map(function (idx) {
        var h = headers[idx];
        // Force the Registration_Type column to the 'entity' role so the frontend
        // shows it as Entity Type (metaFieldType_ would classify it as plain text).
        var type = idx === entityIdx ? 'entity' : metaFieldType_(h);
        return { key: h, type: type, options: Object.keys(optionValues[h] || {}).sort() };
      }),
      extraMetaFields: extraMetaIdx.map(function (idx) { return headers[idx]; }),
      pendingHeader: null,
      eligibilityHeader: eligIdx !== -1 ? headers[eligIdx] : null,
      qualifiedHeader:   qualIdx !== -1 ? headers[qualIdx] : null,
      creditLimitHeader: clIdx   !== -1 ? headers[clIdx]   : null,
      phase0: phase0,
      phase0Buyers: phase0Buyers,
      _isBuyer: true
    };
  } catch (e) {
    return {
      ok: false, error: String(e.message),
      id: 'strideone', name: 'StrideOne',
      entities: [], docs: [], entityColumns: [], entityTypeOptions: [],
      nameHeader: null, entityHeader: null, metaFields: [], extraMetaFields: [],
      pendingHeader: null, eligibilityHeader: null, qualifiedHeader: null, creditLimitHeader: null,
      phase0: 0, phase0Buyers: [], _isBuyer: true
    };
  }
}

function getMbCounts_() {
  try {
    var ss = SpreadsheetApp.openById(MB_SHEET_ID);

    function analyseTab(tabName) {
      try {
        var sh = ss.getSheetByName(tabName);
        if (!sh || sh.getLastRow() < 2) return { count: 0, headers: [], vCol: null, sCol: null, sampleVals: [] };
        var lastCol = sh.getLastColumn();
        var headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
        var vIdx = -1, sIdx = -1;
        for (var c = 0; c < headers.length; c++) {
          var nk = normKey_(headers[c]);
          if (vIdx === -1 && nk.indexOf('vertical') !== -1) vIdx = c;
          if (sIdx === -1 && nk.indexOf('onboarding') !== -1) sIdx = c;
        }
        var mbStatusLower    = MB_STATUS_FILTER.toLowerCase();
        var mbVerticalLower  = MB_VERTICAL_FILTER.toLowerCase();
        var data = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getDisplayValues();
        var uniqueStatus = {};
        data.forEach(function (r) { if (sIdx !== -1) uniqueStatus[String(r[sIdx]).trim()] = true; });
        var count = data.filter(function (r) {
          var verticalOk = vIdx === -1 || String(r[vIdx]).trim().toLowerCase() === mbVerticalLower;
          var statusOk   = sIdx === -1 || String(r[sIdx]).trim().toLowerCase() === mbStatusLower;
          return verticalOk && statusOk;
        }).length;
        return {
          count: count,
          headers: headers,
          vCol: vIdx === -1 ? null : headers[vIdx],
          sCol: sIdx === -1 ? '(not found — status filter skipped)' : headers[sIdx],
          sampleVals: Object.keys(uniqueStatus).slice(0, 20)
        };
      } catch (e) { return { count: 0, headers: [], vCol: null, sCol: null, sampleVals: [], error: String(e.message) }; }
    }

    var sellerInfo = analyseTab('Sellers');
    var buyerInfo  = analyseTab('Buyers');
    return {
      ok:      true,
      sellers: sellerInfo.count,
      buyers:  buyerInfo.count,
      _debug: {
        sellers: { allHeaders: sellerInfo.headers, verticalCol: sellerInfo.vCol, statusCol: sellerInfo.sCol, uniqueStatusVals: sellerInfo.sampleVals },
        buyers:  { allHeaders: buyerInfo.headers,  verticalCol: buyerInfo.vCol,  statusCol: buyerInfo.sCol,  uniqueStatusVals: buyerInfo.sampleVals  }
      }
    };
  } catch (e) {
    return { ok: false, sellers: 0, buyers: 0, error: String(e.message) };
  }
}

/**
 * Debug helper — run this directly in the Apps Script editor (not deployed).
 * Logs the actual column headers and unique values for status/vertical/onboarding
 * columns so you can verify MB_VERTICAL_FILTER and MB_STATUS_FILTER are correct.
 */
function debugMbSheet() {
  var ss = SpreadsheetApp.openById(MB_SHEET_ID);
  ['_mb_sellers', '_mb_buyers'].forEach(function (tabName) {
    var sh = ss.getSheetByName(tabName);
    if (!sh) { Logger.log(tabName + ': TAB NOT FOUND'); return; }
    var lastCol = sh.getLastColumn();
    var lastRow = sh.getLastRow();
    var headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    Logger.log('\n=== ' + tabName + ' === (' + (lastRow - 1) + ' data rows)');
    Logger.log('All headers: ' + JSON.stringify(headers));

    if (lastRow < 2) return;
    var sampleRows = Math.min(lastRow - 1, 200);
    var data = sh.getRange(2, 1, sampleRows, lastCol).getDisplayValues();
    headers.forEach(function (h, i) {
      var nk = normKey_(h);
      if (nk.indexOf('status') !== -1 || nk.indexOf('vertical') !== -1 || nk.indexOf('onboard') !== -1) {
        var uniq = {};
        data.forEach(function (r) { uniq[String(r[i]).trim()] = true; });
        Logger.log('  Col "' + h + '" [' + i + '] unique values: ' + JSON.stringify(Object.keys(uniq)));
      }
    });
  });
}

/* ─────────────────────────── CacheService helpers ─────────────────────── */
// Apps Script CacheService limits each value to 100 KB.
// Larger payloads (e.g. full tab data) are split across numbered chunk keys.

var CACHE_TTL = 300;    // 5 minutes in seconds
var CACHE_VER = 'nbfcv1';

function cacheKey_(name){ return 'nbfc_' + name + '_' + CACHE_VER; }

function cacheSet_(cache, name, obj){
  try {
    var str = JSON.stringify(obj);
    var n = 0, entries = {}, i = 0;
    while (i < str.length){
      entries[cacheKey_(name) + '_' + n++] = str.substring(i, i + 90000);
      i += 90000;
    }
    entries[cacheKey_(name) + '_n'] = String(n);
    cache.putAll(entries, CACHE_TTL);
  } catch(e){ /* ignore write failures */ }
}

function cacheGet_(cache, name){
  try {
    var nStr = cache.get(cacheKey_(name) + '_n');
    if (!nStr) return null;
    var count = parseInt(nStr, 10);
    if (isNaN(count) || count < 1) return null;
    var keys = [];
    for (var n = 0; n < count; n++) keys.push(cacheKey_(name) + '_' + n);
    var map = cache.getAll(keys);
    var result = '';
    for (var j = 0; j < count; j++){
      var chunk = map[cacheKey_(name) + '_' + j];
      if (chunk == null) return null; // partial expiry — treat as full miss
      result += chunk;
    }
    return JSON.parse(result);
  } catch(e){ return null; }
}

function cacheDel_(cache, name){
  try {
    var nStr = cache.get(cacheKey_(name) + '_n');
    var del = [cacheKey_(name) + '_n'];
    if (nStr){
      var c = parseInt(nStr, 10);
      for (var n = 0; n < c; n++) del.push(cacheKey_(name) + '_' + n);
    }
    cache.removeAll(del);
  } catch(e){ /* ignore */ }
}

function clearAllCache_(){
  var cache = CacheService.getScriptCache();
  ['meta', 'tab_billmart', 'tab_capitalxb', 'tab_strideone', 'mb'].forEach(function(k){ cacheDel_(cache, k); });
}

/* ───────────────────────────────── Read API ───────────────────────────────── */

/**
 * Returns all NBFC tab data in one round trip.
 *
 * opts.forceRefresh — when true (sent by the manual Refresh button) bypasses
 * the cache so users always see the latest sheet data.
 *
 * Fast path (all keys in cache): no Sheets API calls → returns in < 200 ms.
 * Slow path (cache miss): reads sheets and repopulates the cache.
 */
function getInitialData(opts) {
  if (opts && opts.forceRefresh) clearAllCache_();

  var cache       = CacheService.getScriptCache();
  var metaC       = cacheGet_(cache, 'meta');
  var billmartC   = cacheGet_(cache, 'tab_billmart');
  var capitalxbC  = cacheGet_(cache, 'tab_capitalxb');
  var strideoneC  = cacheGet_(cache, 'tab_strideone');
  var mbC         = cacheGet_(cache, 'mb');

  // ── Fast path — all data served from cache (no Sheets API calls) ────────
  if (metaC && billmartC && capitalxbC && mbC) {
    // Serve cached StrideOne buyer data if present; otherwise a valid empty
    // buyer object so the tab renders until the slow path repopulates it.
    var strideonefast = strideoneC || {
      ok: true, id: 'strideone', name: 'StrideOne', entities: [], docs: [],
      entityColumns: [], entityTypeOptions: [], nameHeader: null,
      entityHeader: null, metaFields: [], extraMetaFields: [],
      pendingHeader: null, eligibilityHeader: null, qualifiedHeader: null,
      creditLimitHeader: null, phase0: 0, phase0Buyers: [], _isBuyer: true
    };
    return {
      ok: true,
      sheetUrl: metaC.sheetUrl,
      sheetName: metaC.sheetName,
      nbfcs: [billmartC, capitalxbC, strideonefast],
      mbCounts: mbC,
      generatedAt: new Date().toISOString()
    };
  }

  // ── Slow path — read from sheets, populate cache ─────────────────────────
  var ss = getSpreadsheet_();
  var matrix = readRequirementMatrix_(ss);
  var remarksMap = getRemarksMap_(ss);   // GSTIN → { remarks, eta } from Plastic-Remarks

  var billmartData    = getNbfcData_(ss, CONFIG.NBFC_TABS[0], matrix, remarksMap);
  var capitalxbData   = getNbfcData_(ss, CONFIG.NBFC_TABS[1], matrix, remarksMap);
  var strideoneData   = getStrideOneData_(ss, remarksMap);
  var mbData          = getMbCounts_();

  cacheSet_(cache, 'tab_billmart',   billmartData);
  cacheSet_(cache, 'tab_capitalxb',  capitalxbData);
  cacheSet_(cache, 'tab_strideone',  strideoneData);
  cacheSet_(cache, 'mb',             mbData);
  cacheSet_(cache, 'meta', {
    sheetUrl:    ss.getUrl(),
    sheetName:   ss.getName(),
    generatedAt: new Date().toISOString()
  });

  return {
    ok: true,
    sheetUrl: ss.getUrl(),
    sheetName: ss.getName(),
    nbfcs: [billmartData, capitalxbData, strideoneData],
    mbCounts: mbData,
    generatedAt: new Date().toISOString()
  };
}

/* ───────────────────────────────── Write API ──────────────────────────────── */

function toSheetDate_(value) {
  var s = String(value == null ? '' : value).trim();
  if (!s) return '';
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) {
    var d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd-MMM-yyyy');
  }
  return s;
}

/**
 * Creates or updates one entity row in the specified NBFC tab.
 *
 * payload = {
 *   row:          sheet row to update, or null/omit to create,
 *   originalGst:  GST the row had when the form opened (guards re-sort),
 *   meta:         { <header>: value },
 *   statuses:     { <doc header>: 'received' | 'pending' | 'na' },
 *   notes:        { <doc header>: note string }
 * }
 */
function saveEntry(nbfcId, payload) {
  if (!payload || typeof payload !== 'object') throw new Error('Nothing to save.');
  var tabCfg = null;
  CONFIG.NBFC_TABS.forEach(function (t) { if (t.id === nbfcId) tabCfg = t; });
  if (!tabCfg) throw new Error('Unknown NBFC tab id: ' + nbfcId);

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = getSpreadsheet_();
    var sh = findSheet_(ss, tabCfg);
    var layout = readTrackerLayout_(sh, tabCfg.maxCol, tabCfg.docStartCol, tabCfg.docEndCol);
    var matrix = readRequirementMatrix_(ss);

    var meta = payload.meta || {};
    var statuses = payload.statuses || {};
    var notes = payload.notes || {};

    var gstHeader = null, entityHeader = null, nameHeader = null;
    layout.metaIdx.forEach(function (idx) {
      var h = layout.headers[idx];
      var t = metaFieldType_(h);
      if (t === 'gst' && !gstHeader) gstHeader = h;
      if (t === 'entity' && !entityHeader) entityHeader = h;
      if (!nameHeader && normKey_(h).indexOf('name') !== -1) nameHeader = h;
    });

    var name = nameHeader ? String(meta[nameHeader] || '').trim() : '';
    if (nameHeader && !name) throw new Error('Business name is required.');
    var gst = gstHeader ? String(meta[gstHeader] || '').trim().toUpperCase() : '';
    if (gstHeader) meta[gstHeader] = gst;

    var lastRow = sh.getLastRow();
    var lastCol = layout.colCount;
    var existing = lastRow > layout.headerRowNum
      ? sh.getRange(layout.headerRowNum + 1, 1, lastRow - layout.headerRowNum, lastCol).getDisplayValues()
      : [];

    var targetRow = null;
    if (payload.row) {
      var idx0 = Number(payload.row) - (layout.headerRowNum + 1);
      var expected = String(payload.originalGst || '').trim().toUpperCase();
      var gstColIdx = gstHeader ? layout.headers.indexOf(gstHeader) : -1;
      if (idx0 >= 0 && idx0 < existing.length && gstColIdx !== -1 &&
          String(existing[idx0][gstColIdx]).trim().toUpperCase() === expected) {
        targetRow = Number(payload.row);
      } else if (gstColIdx !== -1 && expected) {
        for (var r = 0; r < existing.length; r++) {
          if (String(existing[r][gstColIdx]).trim().toUpperCase() === expected) { targetRow = r + layout.headerRowNum + 1; break; }
        }
      }
      if (!targetRow) {
        // No GST col — fall back to row index
        if (idx0 >= 0 && idx0 < existing.length) targetRow = Number(payload.row);
        else throw new Error('The row being edited was not found (it may have been deleted). Please refresh and try again.');
      }
    } else if (gst && gstHeader) {
      var gCol = layout.headers.indexOf(gstHeader);
      for (var r2 = 0; r2 < existing.length; r2++) {
        if (String(existing[r2][gCol]).trim().toUpperCase() === gst)
          throw new Error('An entity with GST ' + gst + ' already exists. Open it and use Update instead.');
      }
    }

    // Auto-create Annual Turnover column if form submitted one but none exists
    var tvSubmitted = String(meta['Turnover'] || '').trim();
    if (tvSubmitted && !layout.metaIdx.some(function (i) { return metaFieldType_(layout.headers[i]) === 'turnover'; })) {
      sh.insertColumnBefore(layout.pendingIdx + 1);
      sh.getRange(layout.headerRowNum, layout.pendingIdx + 1).setValue('Annual Turnover');
      meta['Annual Turnover'] = tvSubmitted;
      layout = readTrackerLayout_(sh, tabCfg.maxCol, tabCfg.docStartCol, tabCfg.docEndCol);
      lastCol = layout.colCount;
      existing = lastRow > layout.headerRowNum ? sh.getRange(layout.headerRowNum + 1, 1, lastRow - layout.headerRowNum, lastCol).getDisplayValues() : [];
    }

    var entityType  = entityHeader ? String(meta[entityHeader] || '').trim() : '';
    var entityCol   = matchEntityColumn_(matrix, entityType);
    var entityClass = classifyEntityType_(entityType);

    var out = new Array(layout.headers.length);
    for (var c = 0; c < out.length; c++) out[c] = '';

    if (targetRow) {
      var current = sh.getRange(targetRow, 1, 1, lastCol).getDisplayValues()[0];
      for (var c2 = 0; c2 < out.length; c2++) out[c2] = current[c2];
    } else {
      var maxSerial = 0;
      var sIdx = layout.serialIdx;
      existing.forEach(function (row) {
        var n = parseInt(row[sIdx], 10);
        if (!isNaN(n) && n > maxSerial) maxSerial = n;
      });
      out[layout.serialIdx] = maxSerial + 1;
    }

    layout.metaIdx.forEach(function (idx) {
      var h = layout.headers[idx];
      if (!(h in meta)) return;
      var v = String(meta[h] == null ? '' : meta[h]).trim();
      out[idx] = metaFieldType_(h) === 'date' ? toSheetDate_(v) : v;
    });

    var pendingCount = 0;
    layout.docIdx.forEach(function (idx) {
      var h = layout.headers[idx];
      var req = matchRequirementRow_(matrix, h);
      var docNormSE = normKey_(h);
      var alwaysReqSE = ALWAYS_REQUIRED_DOC_FRAGMENTS.some(function(f){ return docNormSE.indexOf(f) !== -1; });
      var applicable;
      if (!isDocApplicableByRules_(h, entityClass)) {
        applicable = false;
      } else if (!alwaysReqSE && req && entityCol) {
        applicable = !!req.requiredBy[entityCol];
      } else {
        applicable = true;
      }
      var status = String(statuses[h] || '').toLowerCase();
      var note = String(notes[h] == null ? '' : notes[h]).trim().replace(/\s+/g, ' ').slice(0, 300);
      var cell;
      if (!applicable) {
        cell = 'NA'; status = 'na';
      } else if (status === 'received') {
        cell = note || 'Received';
      } else if (status === 'na') {
        cell = 'NA';
      } else {
        status = 'pending'; cell = note || 'Pending';
      }
      if (note && status !== 'na' && parseStatus_(cell).status !== status)
        cell = (status === 'received' ? 'Received — ' : 'Pending — ') + note;
      if (status === 'pending') pendingCount++;
      out[idx] = cell;
    });
    out[layout.pendingIdx] = pendingCount;

    var writeRow = targetRow || (lastRow + 1);
    sh.getRange(writeRow, 1, 1, out.length).setValues([out]);

    if (payload.docDriveUrls && typeof payload.docDriveUrls === 'object') {
      Object.keys(payload.docDriveUrls).forEach(function (docKey) {
        var url = String(payload.docDriveUrls[docKey] || '').trim();
        if (!url) return;
        var colIdx = layout.headers.indexOf(docKey);
        if (colIdx >= 0) sh.getRange(writeRow, colIdx + 1).setNote(url);
      });
    }
    if (Array.isArray(payload.docClearUrls)) {
      payload.docClearUrls.forEach(function (docKey) {
        var colIdx = layout.headers.indexOf(docKey);
        if (colIdx >= 0) sh.getRange(writeRow, colIdx + 1).setNote('');
      });
    }

    SpreadsheetApp.flush();
    clearAllCache_();          // ensure the response reflects the just-written data
    var fresh = getInitialData();
    fresh.savedRow = writeRow;
    fresh.savedAction = targetRow ? 'updated' : 'created';
    fresh.savedNbfcId = nbfcId;
    return fresh;
  } finally {
    lock.releaseLock();
  }
}

/* ───────────────────────────── Drive document upload ────────────────────────── */

function getOrCreateEntityFolder_(nbfcName, entityName) {
  var root = DriveApp.getFolderById(DRIVE_ROOT_ID);
  var safe = function (s) { return String(s || '').replace(/[\\\/:\*\?"<>\|]/g, '_').trim() || 'Unknown'; };
  var nbfcFolderName = safe(nbfcName);
  var nbfcIt = root.getFoldersByName(nbfcFolderName);
  var nbfcFolder = nbfcIt.hasNext() ? nbfcIt.next() : root.createFolder(nbfcFolderName);
  var entityFolderName = safe(entityName);
  var entityIt = nbfcFolder.getFoldersByName(entityFolderName);
  return entityIt.hasNext() ? entityIt.next() : nbfcFolder.createFolder(entityFolderName);
}

/**
 * Uploads a document to Drive under DRIVE_ROOT / <NBFC> / <Entity> / <file>.
 *
 * payload = {
 *   nbfcId:      NBFC tab id
 *   row?:        sheet row (1-based) — omit for new entities
 *   entityName?: entity name (for folder/file naming)
 *   gst?:        GST string
 *   docKey:      tracker column header — drives the filename
 *   fileName:    original filename
 *   mimeType:    MIME type
 *   base64Data:  data-URL or raw base64
 * }
 */
function uploadDocument(payload) {
  if (!payload || !payload.base64Data || !payload.docKey)
    throw new Error('uploadDocument: nbfcId, docKey and base64Data are required.');

  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var tabCfg = null;
    CONFIG.NBFC_TABS.forEach(function (t) { if (t.id === payload.nbfcId) tabCfg = t; });
    if (!tabCfg) tabCfg = CONFIG.NBFC_TABS[0]; // fallback

    var entityName = String(payload.entityName || '').trim();
    var docColIdx = -1, sh = null, targetRow = null, layout = null;

    if (payload.row) {
      var ss = getSpreadsheet_();
      sh = findSheet_(ss, tabCfg);
      layout = readTrackerLayout_(sh, tabCfg.maxCol, tabCfg.docStartCol, tabCfg.docEndCol);
      targetRow = Number(payload.row);
      if (targetRow < 2) throw new Error('Invalid row number.');
      var rowData = sh.getRange(targetRow, 1, 1, layout.colCount).getDisplayValues()[0];
      layout.metaIdx.forEach(function (idx) {
        var h = layout.headers[idx];
        if (!entityName && normKey_(h).indexOf('name') !== -1) entityName = String(rowData[idx]).trim();
      });
      docColIdx = layout.headers.indexOf(payload.docKey);
      if (docColIdx < 0) throw new Error('Document column not found: ' + payload.docKey);
    }

    var safe = function (s) { return String(s || '').replace(/[\\\/:\*\?"<>\|]/g, '_').trim(); };
    var origExt = String(payload.fileName || '').split('.').pop();
    var ext = /^[a-zA-Z0-9]{1,8}$/.test(origExt) ? origExt : 'pdf';
    var fname = safe(payload.docKey) + '.' + ext;

    var raw = String(payload.base64Data).replace(/^data:[^;]+;base64,/, '');
    var bytes = Utilities.base64Decode(raw);
    var blob = Utilities.newBlob(bytes, payload.mimeType || 'application/octet-stream', fname);

    var folder = getOrCreateEntityFolder_(tabCfg.name, entityName || String(payload.gst || 'Unknown'));
    var existing = folder.getFilesByName(fname);
    while (existing.hasNext()) existing.next().setTrashed(true);
    var driveFile = folder.createFile(blob);
    driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileUrl = driveFile.getUrl();

    if (sh && targetRow && docColIdx >= 0) {
      sh.getRange(targetRow, docColIdx + 1).setNote(fileUrl);
      SpreadsheetApp.flush();
      clearAllCache_();
      var fresh = getInitialData();
      fresh.savedRow = targetRow;
      fresh.uploadedDoc = payload.docKey;
      fresh.driveUrl = fileUrl;
      fresh.savedNbfcId = payload.nbfcId;
      return fresh;
    }

    return { ok: true, driveUrl: fileUrl, uploadedDoc: payload.docKey };
  } finally {
    lock.releaseLock();
  }
}
