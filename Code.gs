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
var MB_SHEET_ID        = '10RJ1D1GXh-f_7a5M3YMAEt8jDQ7X6jQm2-krTNOBts8';
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

  // Optional: entity-type → document applicability matrix tab.
  REQUIREMENT: { name: 'Seller Requirement', index: -1, signature: ['documents', 'proprietor'] },

  APP_TITLE: 'Recykal · NBFC Document Tracker',
  PROP_BACKEND_ID: 'BACKEND_SHEET_ID'
};

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
    if (docNorm.indexOf(fragment) !== -1 || fragment.indexOf(docNorm) !== -1) {
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
function getNbfcData_(ss, tabCfg, matrix) {
  try {
    var sh = findSheet_(ss, tabCfg);
    var layout = readTrackerLayout_(sh, tabCfg.maxCol, tabCfg.docStartCol, tabCfg.docEndCol);

    var nameHeaderKey = null, entityHeaderKey = null;
    layout.metaIdx.forEach(function (idx) {
      var h = layout.headers[idx];
      if (!nameHeaderKey && normKey_(h).indexOf('name') !== -1) nameHeaderKey = h;
      if (!entityHeaderKey && metaFieldType_(h) === 'entity') entityHeaderKey = h;
    });

    // Build doc list — match each column against the requirement matrix
    var docs = layout.docIdx.map(function (idx) {
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
          completion: applicable ? Math.round((received / applicable) * 1000) / 10 : 0
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
      pendingHeader: layout.headers[layout.pendingIdx]
    };
  } catch (e) {
    return {
      ok: false, error: String(e.message),
      id: tabCfg.id, name: tabCfg.name,
      entities: [], docs: [], entityColumns: [], entityTypeOptions: [],
      nameHeader: null, entityHeader: null, metaFields: []
    };
  }
}

/* ──────────────────── Marketplace counts (Open Marketplace · Onboarded) ─── */

var MB_STATUS_FILTER = 'Onboarded';

/**
 * Counts rows in a marketplace tab where:
 *   Vertical = "Open Marketplace"  AND  Status = "Onboarded"
 * Header detection is case-insensitive; both columns must be present.
 */
function getMbCounts_() {
  try {
    var ss = SpreadsheetApp.openById(MB_SHEET_ID);

    function countOmpRows(tabName) {
      try {
        var sh = ss.getSheetByName(tabName);
        if (!sh || sh.getLastRow() < 2) return 0;
        var lastCol = sh.getLastColumn();
        var headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
        var vIdx = -1, sIdx = -1;
        for (var c = 0; c < headers.length; c++) {
          var nk = normKey_(headers[c]);
          if (vIdx === -1 && nk.indexOf('vertical') !== -1) vIdx = c;
          if (sIdx === -1 && nk === 'status')               sIdx = c;
        }
        var data = sh.getRange(2, 1, sh.getLastRow() - 1, lastCol).getDisplayValues();
        return data.filter(function (r) {
          var verticalOk = vIdx === -1 || String(r[vIdx]).trim() === MB_VERTICAL_FILTER;
          var statusOk   = sIdx === -1 || String(r[sIdx]).trim() === MB_STATUS_FILTER;
          return verticalOk && statusOk;
        }).length;
      } catch (e) { return 0; }
    }

    return {
      ok: true,
      sellers: countOmpRows('_mb_sellers'),
      buyers:  countOmpRows('_mb_buyers')
    };
  } catch (e) {
    return { ok: false, sellers: 0, buyers: 0, error: String(e.message) };
  }
}

/* ───────────────────────────────── Read API ───────────────────────────────── */

/**
 * Returns all NBFC tab data in one round trip.
 */
function getInitialData() {
  var ss = getSpreadsheet_();
  var matrix = readRequirementMatrix_(ss); // gracefully returns empty if tab absent

  var nbfcs = CONFIG.NBFC_TABS.map(function (tab) {
    // StrideOne tracks buyers — integration with the buyer sheet is not yet active.
    if (tab.id === 'strideone') {
      return {
        ok: true, id: tab.id, name: tab.name,
        entities: [], docs: [], entityColumns: [], entityTypeOptions: [],
        nameHeader: null, entityHeader: null, metaFields: [], extraMetaFields: [],
        pendingHeader: null, _buyerPlaceholder: true
      };
    }
    return getNbfcData_(ss, tab, matrix);
  });

  return {
    ok: true,
    sheetUrl: ss.getUrl(),
    sheetName: ss.getName(),
    nbfcs: nbfcs,
    mbCounts: getMbCounts_(),
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
