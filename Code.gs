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

var CONFIG = {
  // Native Google Sheet ID (already confirmed native; no xlsx conversion needed).
  SOURCE_FILE_ID: '1RoHWbZyHhNKlweWXD4AMSZfB5ONdktPcVayOkpPgjpo',

  // One entry per NBFC tracker tab.
  // maxCol caps how far right the tab is read so scratch columns to the right
  // never pollute KPIs, cards, the matrix, or form saves.
  NBFC_TABS: [
    { id: 'billmart',  name: 'Billmart',   maxCol: 40 },
    { id: 'capitalxb', name: 'Capital XB', maxCol: 40 },
    { id: 'strideone', name: 'StrideOne',  maxCol: 40 },
  ],

  // Optional: entity-type → document applicability matrix tab.
  REQUIREMENT: { name: 'Seller Requirement', index: -1, signature: ['documents', 'proprietor'] },

  APP_TITLE: 'Recykal · NBFC Document Tracker',
  PROP_BACKEND_ID: 'BACKEND_SHEET_ID'
};

/* ─────────────────────────── Web-app entry ─────────────────────────────── */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(CONFIG.APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/* ───────────────────────── Spreadsheet access ──────────────────────────── */

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

/* ─────────────────────────── Tab discovery ─────────────────────────────── */

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

/* ─────────────────────────── Status vocab ──────────────────────────────── */

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

/* ──────────────────── Requirement matrix (optional tab) ─────────────────── */

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

/* ─────────────────────────── Tracker layout ─────────────────────────────── */

function readTrackerLayout_(sh, maxCol) {
  var lastCol = sh.getLastColumn();
  if (maxCol && maxCol > 0 && maxCol < lastCol) lastCol = maxCol;
  var headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0]
    .map(function (h) { return String(h).trim(); });
  var pendingIdx = -1, i;
  for (i = 0; i < headers.length; i++) {
    if (normKey_(headers[i]) === 'pendingdocument' || normKey_(headers[i]) === 'pendingdocuments') {
      pendingIdx = i; break;
    }
  }
  if (pendingIdx === -1)
    throw new Error('The "' + sh.getName() + '" tab has no "Pending Document" column within the first ' + lastCol + ' columns.');
  var serialIdx = 0;
  var metaIdx = [];
  for (i = 0; i < pendingIdx; i++) {
    if (i === serialIdx) continue;
    if (headers[i]) metaIdx.push(i);
  }
  var docIdx = [];
  for (i = pendingIdx + 1; i < headers.length; i++) {
    if (headers[i]) docIdx.push(i);
  }
  return { headers: headers, colCount: lastCol, serialIdx: serialIdx, metaIdx: metaIdx, pendingIdx: pendingIdx, docIdx: docIdx };
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

/* ────────────────────────── Per-NBFC data reader ───────────────────────── */

/**
 * Reads one NBFC tracker tab and returns all entity rows with parsed doc statuses.
 * On any error, returns an error object so getInitialData() can still succeed.
 */
function getNbfcData_(ss, tabCfg, matrix) {
  try {
    var sh = findSheet_(ss, tabCfg);
    var layout = readTrackerLayout_(sh, tabCfg.maxCol);

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

    if (lastRow > 1) {
      var dataRange = sh.getRange(2, 1, lastRow - 1, layout.colCount);
      var values = dataRange.getDisplayValues();
      var allNotes = dataRange.getNotes();

      values.forEach(function (row, i) {
        var meta = {};
        layout.metaIdx.forEach(function (idx) { meta[layout.headers[idx]] = String(row[idx]).trim(); });
        var hasIdentity = layout.metaIdx.some(function (idx) { return String(row[idx]).trim() !== ''; });
        if (!hasIdentity) return;

        layout.metaIdx.forEach(function (idx) {
          var h = layout.headers[idx], v = String(row[idx]).trim();
          if (!v) return;
          (optionValues[h] = optionValues[h] || {})[v] = true;
        });

        var serial = parseInt(row[layout.serialIdx], 10);
        var entityType = entityHeaderKey ? (meta[entityHeaderKey] || '') : '';
        var entityCol = matchEntityColumn_(matrix, entityType);
        var rowNotes = allNotes[i] || [];
        var docStates = {};
        var received = 0, pending = 0, na = 0;

        layout.docIdx.forEach(function (idx) {
          var h = layout.headers[idx];
          var req = matchRequirementRow_(matrix, h);
          var applicable = (!req || !entityCol) ? true : !!req.requiredBy[entityCol];
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
          row: i + 2,
          serial: isNaN(serial) ? '' : serial,
          meta: meta,
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

/* ───────────────────────────── Read API ────────────────────────────────── */

/**
 * Returns all NBFC tab data in one round trip.
 */
function getInitialData() {
  var ss = getSpreadsheet_();
  var matrix = readRequirementMatrix_(ss); // gracefully returns empty if tab absent

  var nbfcs = CONFIG.NBFC_TABS.map(function (tab) {
    return getNbfcData_(ss, tab, matrix);
  });

  return {
    ok: true,
    sheetUrl: ss.getUrl(),
    sheetName: ss.getName(),
    nbfcs: nbfcs,
    generatedAt: new Date().toISOString()
  };
}

/* ───────────────────────────── Write API ───────────────────────────────── */

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
    var layout = readTrackerLayout_(sh, tabCfg.maxCol);
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
    var existing = lastRow > 1
      ? sh.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues()
      : [];

    var targetRow = null;
    if (payload.row) {
      var idx0 = Number(payload.row) - 2;
      var expected = String(payload.originalGst || '').trim().toUpperCase();
      var gstColIdx = gstHeader ? layout.headers.indexOf(gstHeader) : -1;
      if (idx0 >= 0 && idx0 < existing.length && gstColIdx !== -1 &&
          String(existing[idx0][gstColIdx]).trim().toUpperCase() === expected) {
        targetRow = Number(payload.row);
      } else if (gstColIdx !== -1 && expected) {
        for (var r = 0; r < existing.length; r++) {
          if (String(existing[r][gstColIdx]).trim().toUpperCase() === expected) { targetRow = r + 2; break; }
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
      sh.getRange(1, layout.pendingIdx + 1).setValue('Annual Turnover');
      meta['Annual Turnover'] = tvSubmitted;
      layout = readTrackerLayout_(sh, tabCfg.maxCol);
      lastCol = layout.colCount;
      existing = lastRow > 1 ? sh.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues() : [];
    }

    var entityType = entityHeader ? String(meta[entityHeader] || '').trim() : '';
    var entityCol = matchEntityColumn_(matrix, entityType);

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
      var applicable = (!req || !entityCol) ? true : !!req.requiredBy[entityCol];
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

/* ──────────────────────── Drive document upload ────────────────────────── */

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
      layout = readTrackerLayout_(sh, tabCfg.maxCol);
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
