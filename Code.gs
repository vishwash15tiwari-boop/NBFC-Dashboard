/**
 * ─────────────────────────────────────────────────────────────────────────────
 *  Recykal · Seller Onboarding — NBFC Document Tracker (Web App backend)
 * ─────────────────────────────────────────────────────────────────────────────
 *  Backend for the one-page executive dashboard + seller entry form served by
 *  Index.html. The spreadsheet is the single source of truth:
 *
 *    • "Seller_NBFC Tracker"  (4th tab)  — primary database, one row per seller
 *    • "Seller Requirement"   (3rd tab)  — mandatory document checklist by
 *                                          business category (read dynamically;
 *                                          document names are NEVER hardcoded)
 *
 *  The sheet's structure is preserved exactly: this script only appends new
 *  seller rows or updates existing ones, writing values in the sheet's own
 *  column order and status vocabulary (Received / Pending / NA / free-text
 *  notes).
 *
 *  SOURCE_FILE_ID points at the native Google Sheet "NBFC Document Tracker"
 *  and is used directly. (Safety net: if the ID is ever swapped for an
 *  uploaded .xlsx — which Apps Script cannot read or write in place — the
 *  script converts it once into a native Google Sheet ("… (Live)") in the
 *  same Drive folder, stores the new ID in Script Properties, and uses that
 *  as the live backend from then on.)
 * ─────────────────────────────────────────────────────────────────────────────
 */

var CONFIG = {
  // Drive file ID of the workbook (native Google Sheet, or xlsx upload).
  SOURCE_FILE_ID: '1RoHWbZyHhNKlweWXD4AMSZfB5ONdktPcVayOkpPgjpo',

  // Tab discovery: matched by name first, then by header signature, then by
  // position (0-based index) as a last resort.
  TRACKER: { name: 'Seller_NBFC Tracker', index: 3, signature: ['seller business name', 'pending document'] },
  REQUIREMENT: { name: 'Seller Requirement', index: 2, signature: ['documents', 'proprietor'] },

  APP_TITLE: 'Recykal · NBFC Document Tracker',
  PROP_BACKEND_ID: 'BACKEND_SHEET_ID'
};

/* ───────────────────────────── Web app entry ───────────────────────────── */

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(CONFIG.APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/* ─────────────────────────── Backend resolution ────────────────────────── */

/**
 * Returns the live native Google Sheet, converting the source .xlsx once if
 * needed. The converted sheet keeps every tab, header and value unchanged.
 */
function getSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();
  var savedId = props.getProperty(CONFIG.PROP_BACKEND_ID);
  if (savedId) {
    try { return SpreadsheetApp.openById(savedId); } catch (e) { /* stale — re-resolve below */ }
  }
  try {
    var direct = SpreadsheetApp.openById(CONFIG.SOURCE_FILE_ID); // already a native Sheet
    props.setProperty(CONFIG.PROP_BACKEND_ID, direct.getId());
    return direct;
  } catch (e) { /* not native — convert once */ }
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    savedId = props.getProperty(CONFIG.PROP_BACKEND_ID); // another request may have converted meanwhile
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

/** One-time multipart upload to Drive that re-imports the xlsx as a native Sheet. */
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
    {
      method: 'post',
      contentType: 'multipart/related; boundary=' + boundary,
      payload: payload,
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    }
  );
  if (res.getResponseCode() >= 300) {
    throw new Error('Could not convert the Excel workbook to a native Google Sheet. ' +
      'Drive API said: ' + res.getContentText());
  }
  return JSON.parse(res.getContentText()).id;
}

/* ───────────────────────────── Tab discovery ───────────────────────────── */

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
  for (i = 0; i < sheets.length; i++) {
    sh = sheets[i];
    if (sh.getLastRow() < 1 || sh.getLastColumn() < 1) continue;
    var header = sh.getRange(1, 1, 1, sh.getLastColumn()).getDisplayValues()[0]
      .map(function (h) { return String(h).toLowerCase(); }).join('|');
    var hit = spec.signature.every(function (tok) { return header.indexOf(tok) !== -1; });
    if (hit) return sh;
  }
  if (spec.index >= 0 && spec.index < sheets.length) return sheets[spec.index];
  throw new Error('Could not locate the "' + spec.name + '" tab in the spreadsheet. ' +
    'Please make sure the tab exists and its header row is intact.');
}

/* ─────────────────────────── Status vocabulary ─────────────────────────── */

/**
 * Interprets a tracker cell exactly the way the sheet already uses it:
 *   "Received"                          → received
 *   "NA" / "N/A" / "Not Applicable"     → na (not required for this entity)
 *   "Pending" or any note that contains
 *   the word "pending"                  → pending (note preserved)
 *   blank / "-" / other free text       → pending (treated as an open item)
 */
function parseStatus_(raw) {
  var s = String(raw == null ? '' : raw).trim();
  var l = s.toLowerCase();
  if (!s || s === '-') return { status: 'pending', note: '' };
  if (l === 'na' || l === 'n/a' || l === 'not applicable') return { status: 'na', note: '' };
  if (l.indexOf('pending') !== -1) {
    return { status: 'pending', note: l === 'pending' ? '' : s };
  }
  if (l === 'received' || l === 'yes' || l === 'done' || l === 'submitted' ||
      l.indexOf('receiv') !== -1 || l.indexOf('provided') !== -1) {
    return { status: 'received', note: (l === 'received' || l === 'yes') ? '' : s };
  }
  return { status: 'pending', note: s };
}

/* ─────────────────────── Requirement matrix (3rd tab) ──────────────────── */

/**
 * Reads the Seller Requirement tab dynamically:
 *   header:  Sr. No. | Documents | <entity column> | <entity column> | …
 *   rows:    a checkmark (✔ / ✓ / yes) marks the document mandatory for that
 *            business category; "-" or blank means not applicable.
 */
function readRequirementMatrix_(ss) {
  var sh = findSheet_(ss, CONFIG.REQUIREMENT);
  var values = sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getDisplayValues();
  var headerRow = -1, docCol = -1, r, c;
  for (r = 0; r < values.length && headerRow === -1; r++) {
    for (c = 0; c < values[r].length; c++) {
      if (normKey_(values[r][c]) === 'documents') { headerRow = r; docCol = c; break; }
    }
  }
  if (headerRow === -1) {
    throw new Error('The "Seller Requirement" tab has no "Documents" header — cannot read the checklist.');
  }
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
}

/** Fuzzy-matches a requirement row to a tracker document column (normalized containment). */
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

/**
 * Maps a stored entity type (e.g. "Partnership Firm", "Private Limited
 * Company") to the matching requirement column (e.g. "Partnership",
 * "Private Limited"), preferring the longest label so "Private Limited
 * Company" resolves to "Private Limited" and not "Limited".
 */
function matchEntityColumn_(matrix, entityType) {
  var e = normKey_(entityType);
  if (!e) return null;
  var best = null, bestLen = 0;
  matrix.entityColumns.forEach(function (label) {
    var n = normKey_(label);
    if ((e.indexOf(n) !== -1 || n.indexOf(e) !== -1) && n.length > bestLen) {
      best = label; bestLen = n.length;
    }
  });
  return best;
}

/* ───────────────────── Tracker structure (4th tab) ─────────────────────── */

/**
 * Reads the tracker header row and splits it, entirely by position, into:
 *   serial column ("S.No") · meta columns (Region … Entity Type) ·
 *   the "Pending Document" counter · document columns (everything after it).
 * Nothing about the structure is assumed beyond the presence of the
 * "Pending Document" header the sheet already has.
 */
function readTrackerLayout_(sh) {
  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(1, 1, 1, lastCol).getDisplayValues()[0]
    .map(function (h) { return String(h).trim(); });
  var pendingIdx = -1, i;
  for (i = 0; i < headers.length; i++) {
    if (normKey_(headers[i]) === 'pendingdocument' || normKey_(headers[i]) === 'pendingdocuments') {
      pendingIdx = i; break;
    }
  }
  if (pendingIdx === -1) {
    throw new Error('The tracker tab has no "Pending Document" column — the sheet structure has changed.');
  }
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
  return { headers: headers, serialIdx: serialIdx, metaIdx: metaIdx, pendingIdx: pendingIdx, docIdx: docIdx };
}

/** Detects an input affordance for a meta column from its header name only. */
function metaFieldType_(header) {
  var n = normKey_(header);
  if (n.indexOf('gst') !== -1) return 'gst';
  if (n.indexOf('date') !== -1) return 'date';
  if (n.indexOf('entitytype') !== -1 || n.indexOf('businesstype') !== -1) return 'entity';
  return 'text';
}

/* ────────────────────────────── Read API ───────────────────────────────── */

/**
 * Everything the front end needs, in one round trip:
 * layout, requirement matrix, per-document applicability, and every seller
 * row with parsed statuses and completion aggregates.
 */
function getInitialData() {
  var ss = getSpreadsheet_();
  var tracker = findSheet_(ss, CONFIG.TRACKER);
  var layout = readTrackerLayout_(tracker);
  var matrix = readRequirementMatrix_(ss);

  var docs = layout.docIdx.map(function (idx) {
    var header = layout.headers[idx];
    var req = matchRequirementRow_(matrix, header);
    var requiredBy = {};
    matrix.entityColumns.forEach(function (label) {
      requiredBy[label] = req ? !!req.requiredBy[label] : true; // unmatched docs default to required
    });
    return { key: header, requiredBy: requiredBy };
  });

  var lastRow = tracker.getLastRow();
  var sellers = [];
  var maxSerial = 0;
  var optionValues = {}; // distinct existing values per meta header, for form suggestions

  if (lastRow > 1) {
    var dataRange = tracker.getRange(2, 1, lastRow - 1, tracker.getLastColumn());
    var values = dataRange.getDisplayValues();
    var allNotes = dataRange.getNotes();
    values.forEach(function (row, i) {
      var meta = {};
      layout.metaIdx.forEach(function (idx) { meta[layout.headers[idx]] = String(row[idx]).trim(); });
      var hasIdentity = layout.metaIdx.some(function (idx) { return String(row[idx]).trim() !== ''; });
      if (!hasIdentity) return; // skip fully blank rows

      layout.metaIdx.forEach(function (idx) {
        var h = layout.headers[idx], v = String(row[idx]).trim();
        if (!v) return;
        (optionValues[h] = optionValues[h] || {})[v] = true;
      });

      var serial = parseInt(row[layout.serialIdx], 10);
      if (!isNaN(serial) && serial > maxSerial) maxSerial = serial;

      var rowNotes = allNotes[i] || [];
      var docStates = {};
      var received = 0, pending = 0, na = 0;
      layout.docIdx.forEach(function (idx) {
        var parsed = parseStatus_(row[idx]);
        var cellNote = String(rowNotes[idx] || '').trim();
        var driveUrl = /^https?:\/\/\S+$/.test(cellNote) ? cellNote : '';
        docStates[layout.headers[idx]] = { raw: String(row[idx]).trim(), status: parsed.status, note: parsed.note, driveUrl: driveUrl };
        if (parsed.status === 'received') received++;
        else if (parsed.status === 'na') na++;
        else pending++;
      });
      var applicable = received + pending;
      sellers.push({
        row: i + 2,                       // 1-based sheet row
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

  // Build entity options from actual sheet values first, then fall back to matrix column
  // labels only for entity types not already represented (avoids "Partnership" + "Partnership Firm").
  var entityOptions = {};
  sellers.forEach(function (s) {
    layout.metaIdx.forEach(function (idx) {
      var h = layout.headers[idx];
      if (metaFieldType_(h) === 'entity' && s.meta[h]) entityOptions[s.meta[h]] = true;
    });
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
    sheetUrl: ss.getUrl(),
    sheetName: ss.getName(),
    trackerName: tracker.getName(),
    metaFields: layout.metaIdx.map(function (idx) {
      var h = layout.headers[idx];
      return {
        key: h,
        type: metaFieldType_(h),
        options: Object.keys(optionValues[h] || {}).sort()
      };
    }),
    pendingHeader: layout.headers[layout.pendingIdx],
    docs: docs,
    entityColumns: matrix.entityColumns,
    entityTypeOptions: Object.keys(entityOptions).sort(),
    sellers: sellers,
    generatedAt: new Date().toISOString()
  };
}

/* ────────────────────────────── Write API ──────────────────────────────── */

/** Formats a date for the sheet the way existing rows store it: 01-Feb-2018. */
function toSheetDate_(value) {
  var s = String(value == null ? '' : value).trim();
  if (!s) return '';
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/); // ISO from the <input type="date">
  if (m) {
    var d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd-MMM-yyyy');
  }
  return s; // already in the sheet's own format — preserve as-is
}

/**
 * Creates or updates one seller row, writing values in the sheet's exact
 * column order and vocabulary. Never touches any other row or column.
 *
 * payload = {
 *   row:        sheet row number to update, or null to create,
 *   originalGst: GST the row had when the form was opened (guards against
 *                the sheet being re-sorted while the form was open),
 *   meta:       { <meta header>: value },
 *   statuses:   { <doc header>: 'received' | 'pending' | 'na' },
 *   notes:      { <doc header>: optional free-text note }
 * }
 */
function saveSeller(payload) {
  if (!payload || typeof payload !== 'object') throw new Error('Nothing to save.');
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var ss = getSpreadsheet_();
    var tracker = findSheet_(ss, CONFIG.TRACKER);
    var layout = readTrackerLayout_(tracker);
    var matrix = readRequirementMatrix_(ss);

    var meta = payload.meta || {};
    var statuses = payload.statuses || {};
    var notes = payload.notes || {};

    // Identify the GST + entity + name columns from the layout (never by letter).
    var gstHeader = null, entityHeader = null, nameHeader = null;
    layout.metaIdx.forEach(function (idx) {
      var h = layout.headers[idx];
      var t = metaFieldType_(h);
      if (t === 'gst' && !gstHeader) gstHeader = h;
      if (t === 'entity' && !entityHeader) entityHeader = h;
      if (!nameHeader && normKey_(h).indexOf('name') !== -1) nameHeader = h;
    });

    var name = nameHeader ? String(meta[nameHeader] || '').trim() : '';
    if (nameHeader && !name) throw new Error('Seller business name is required.');
    var gst = gstHeader ? String(meta[gstHeader] || '').trim().toUpperCase() : '';
    if (gstHeader) meta[gstHeader] = gst;

    var lastRow = tracker.getLastRow();
    var lastCol = tracker.getLastColumn();
    var existing = lastRow > 1
      ? tracker.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues()
      : [];

    // Resolve the target row.
    var targetRow = null; // 1-based sheet row
    if (payload.row) {
      var idx0 = Number(payload.row) - 2;
      var expected = String(payload.originalGst || '').trim().toUpperCase();
      var gstColIdx = gstHeader ? layout.headers.indexOf(gstHeader) : -1;
      if (idx0 >= 0 && idx0 < existing.length && gstColIdx !== -1 &&
          String(existing[idx0][gstColIdx]).trim().toUpperCase() === expected) {
        targetRow = Number(payload.row);
      } else if (gstColIdx !== -1 && expected) {
        for (var r = 0; r < existing.length; r++) { // row moved — find it by its original GST
          if (String(existing[r][gstColIdx]).trim().toUpperCase() === expected) { targetRow = r + 2; break; }
        }
      }
      if (!targetRow) {
        throw new Error('The row being edited was not found in the sheet (it may have been deleted). ' +
          'Please refresh and try again.');
      }
    } else if (gst && gstHeader) {
      var gCol = layout.headers.indexOf(gstHeader);
      for (var r2 = 0; r2 < existing.length; r2++) {
        if (String(existing[r2][gCol]).trim().toUpperCase() === gst) {
          throw new Error('A seller with GST ' + gst + ' already exists (' +
            String(existing[r2][layout.headers.indexOf(nameHeader)] || 'row ' + (r2 + 2)).trim() +
            '). Open that seller and use Update instead.');
        }
      }
    }

    // Applicability from the requirement matrix for this seller's entity type.
    var entityType = entityHeader ? String(meta[entityHeader] || '').trim() : '';
    var entityCol = matchEntityColumn_(matrix, entityType);

    // Compose the full row in exact column order.
    var out = new Array(layout.headers.length);
    for (var c = 0; c < out.length; c++) out[c] = '';

    if (targetRow) {
      var current = tracker.getRange(targetRow, 1, 1, lastCol).getDisplayValues()[0];
      for (var c2 = 0; c2 < out.length; c2++) out[c2] = current[c2]; // start from what's there
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
      if (!(h in meta)) return; // untouched fields keep their current value
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
      } else { // pending (default when the form sends nothing for a column)
        status = 'pending';
        cell = note || 'Pending';
      }
      // A note must still parse back to the status it was saved with.
      if (note && status !== 'na' && parseStatus_(cell).status !== status) {
        cell = note + (status === 'received' ? ' — Received' : ' — Pending');
      }
      if (status === 'pending') pendingCount++;
      out[idx] = cell;
    });
    out[layout.pendingIdx] = pendingCount;

    var writeRow = targetRow || (lastRow + 1);
    tracker.getRange(writeRow, 1, 1, out.length).setValues([out]);

    // Persist per-document Drive URLs into cell notes.
    if (payload.docDriveUrls && typeof payload.docDriveUrls === 'object') {
      Object.keys(payload.docDriveUrls).forEach(function (docKey) {
        var url = String(payload.docDriveUrls[docKey] || '').trim();
        if (!url) return;
        var colIdx = layout.headers.indexOf(docKey);
        if (colIdx >= 0) tracker.getRange(writeRow, colIdx + 1).setNote(url);
      });
    }
    // Clear notes for documents whose Drive URL was removed in this session.
    if (Array.isArray(payload.docClearUrls)) {
      payload.docClearUrls.forEach(function (docKey) {
        var colIdx = layout.headers.indexOf(docKey);
        if (colIdx >= 0) tracker.getRange(writeRow, colIdx + 1).setNote('');
      });
    }

    SpreadsheetApp.flush();

    var fresh = getInitialData();
    fresh.savedRow = writeRow;
    fresh.savedAction = targetRow ? 'updated' : 'created';
    return fresh;
  } finally {
    lock.releaseLock();
  }
}

/* ────────────────────────── Drive integration ───────────────────────────── */

var DRIVE_ROOT_ID = '1i5melXCocWrV9rR-3gM75wwSwWy7Dqit';

function getOrCreateSellerFolder_(sellerName, gst) {
  var root = DriveApp.getFolderById(DRIVE_ROOT_ID);
  var safe = function (s) { return String(s || '').replace(/[\\\/:\*\?"<>\|]/g, '_').trim(); };
  var folderName = sellerName
    ? safe(sellerName)
    : safe(gst || 'Unknown');
  var it = root.getFoldersByName(folderName);
  if (it.hasNext()) return it.next();
  return root.createFolder(folderName);
}

/**
 * Uploads a single document file to the seller's Drive subfolder, named after the
 * document type (docKey). Works for existing sellers (pass row — cell note updated
 * immediately) and new sellers not yet saved (pass sellerName + gst — URL returned
 * for saveSeller to persist later).
 *
 * payload = {
 *   row?:        sheet row number (1-based) — omit for new sellers
 *   sellerName?: seller name for folder/file naming
 *   gst?:        GST string (used to resolve name from sheet when row is given)
 *   docKey:      tracker column header — drives the filename
 *   fileName:    original filename (extension extracted for the Drive file)
 *   mimeType:    MIME type string
 *   base64Data:  data-URL string (data:[type];base64,[data]) or raw base64
 * }
 */
function uploadDocument(payload) {
  if (!payload || !payload.base64Data || !payload.docKey) {
    throw new Error('uploadDocument: docKey and base64Data are required.');
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var sellerName = String(payload.sellerName || '').trim();
    var docColIdx = -1, tracker = null, targetRow = null, layout = null;

    if (payload.row) {
      var ss = getSpreadsheet_();
      tracker = findSheet_(ss, CONFIG.TRACKER);
      layout = readTrackerLayout_(tracker);
      targetRow = Number(payload.row);
      if (targetRow < 2) throw new Error('Invalid row number.');
      var rowData = tracker.getRange(targetRow, 1, 1, tracker.getLastColumn()).getDisplayValues()[0];
      layout.metaIdx.forEach(function (idx) {
        var h = layout.headers[idx];
        if (!sellerName && normKey_(h).indexOf('name') !== -1) sellerName = String(rowData[idx]).trim();
      });
      docColIdx = layout.headers.indexOf(payload.docKey);
      if (docColIdx < 0) throw new Error('Document column not found: ' + payload.docKey);
    }

    // File named as [safe docKey].[original extension].
    var safe = function (s) { return String(s || '').replace(/[\\\/:\*\?"<>\|]/g, '_').trim(); };
    var origExt = String(payload.fileName || '').split('.').pop();
    var ext = /^[a-zA-Z0-9]{1,8}$/.test(origExt) ? origExt : 'pdf';
    var fname = safe(payload.docKey) + '.' + ext;

    var raw = String(payload.base64Data).replace(/^data:[^;]+;base64,/, '');
    var bytes = Utilities.base64Decode(raw);
    var blob = Utilities.newBlob(bytes, payload.mimeType || 'application/octet-stream', fname);

    var folder = getOrCreateSellerFolder_(sellerName, String(payload.gst || '').trim());
    var existing = folder.getFilesByName(fname);
    while (existing.hasNext()) existing.next().setTrashed(true);
    var driveFile = folder.createFile(blob);
    driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var fileUrl = driveFile.getUrl();

    // For existing sellers: store URL in the doc cell's note immediately.
    if (tracker && targetRow && docColIdx >= 0) {
      tracker.getRange(targetRow, docColIdx + 1).setNote(fileUrl);
      SpreadsheetApp.flush();
      var fresh = getInitialData();
      fresh.savedRow = targetRow;
      fresh.uploadedDoc = payload.docKey;
      fresh.driveUrl = fileUrl;
      return fresh;
    }

    // For new sellers: return just the URL; saveSeller will write the note.
    return { ok: true, driveUrl: fileUrl, uploadedDoc: payload.docKey };
  } finally {
    lock.releaseLock();
  }
}
