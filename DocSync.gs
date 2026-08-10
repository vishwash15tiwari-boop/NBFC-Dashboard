/**
 * ─────────────────────────────────────────────────────────────────────────────
 *  DocSync — Source ➜ Destination document-column synchroniser
 * ─────────────────────────────────────────────────────────────────────────────
 *  Copies document-related columns from the NBFC tracker (source) into the
 *  master sheet (destination), keyed solely on GSTIN.
 *
 *  Design rules honoured throughout:
 *    • GSTIN is the only join key, compared trimmed + upper-cased.
 *    • Nothing is hardcoded: tab names, header positions and document names are
 *      all discovered at runtime from row 1 of each sheet.
 *    • Only document columns are written. Every other column — and every
 *      destination cell holding a formula — is left exactly as it was, so
 *      formatting, validations, filters, remarks and manual inputs survive.
 *    • All I/O is batched: one read per sheet, one write per contiguous block
 *      of target columns. There are no per-cell reads or writes.
 *    • Values are copied verbatim. Source formulas (including =HYPERLINK) are
 *      carried across as formulas, and rich-text links are preserved.
 *
 *  Entry points:
 *    syncDocuments()            — run the sync (safe to run manually or on a trigger)
 *    installDocSyncTrigger()    — install the recurring trigger (idempotent)
 *    removeDocSyncTriggers()    — remove it again
 *    syncMetabaseThenDocs()     — ordered run: Metabase refresh, then doc sync
 *    previewDocSyncPlan()       — dry run: log what WOULD change, write nothing
 *
 *  ⚠ INTERACTION WITH MetabaseSync.gs
 *  MetabaseSync rewrites its managed tabs by clearing rows 2..lastRow across
 *  columns 1..headers.length, on a one-minute trigger, and it populates document
 *  columns itself. Any value this script writes INSIDE that managed block will be
 *  erased on the next Metabase run. The sync therefore detects that overlap and
 *  reports it as a warning on every run rather than failing silently. To resolve
 *  it, pick one: (a) drive both from syncMetabaseThenDocs() and drop the
 *  standalone one-minute trigger, or (b) remove those document columns from
 *  MetabaseSync's COLUMN_ORDER/DOC_COLUMNS so this script becomes their sole
 *  owner. Columns this script creates itself sit beyond the managed width and
 *  are unaffected either way.
 * ─────────────────────────────────────────────────────────────────────────────
 */

var DOCSYNC_CFG = {
  SOURCE_ID: '1RoHWbZyHhNKlweWXD4AMSZfB5ONdktPcVayOkpPgjpo',
  DEST_ID:   '1d57KGl00-pGWVjYKouyMu8jt0Y4UMEc2HaHWtMWPjeM',

  /* Header aliases that identify the GSTIN key column, matched case- and
     punctuation-insensitively. Extend rather than rename, so older tabs keep working. */
  KEY_ALIASES: ['gstin', 'gst', 'gst number', 'gst no', 'gstin number',
                'seller gstin', 'buyer gstin', 'seller gst number', 'buyer gst number',
                'seller_gst_number', 'buyer_gst_number'],

  /* Identity / operational headers that are never treated as document columns,
     even when they appear in both sheets. Compared on the normalised header. */
  NON_DOC_HEADERS: [
    'no', 'sr no', 's no', 'serial', 'row',
    'seller business name', 'buyer business name', 'business name', 'entity name', 'name',
    'poc name', 'contact person name', 'company name', 'organization',
    'region', 'vertical', 'state', 'city', 'district', 'pincode', 'address',
    'seller type', 'buyer type', 'business category', 'customer type', 'entity type',
    'vintage with recykal', 'vintage', 'business vintage', 'effective date of registration',
    'onboarding status', 'status', 'seller status', 'active',
    'mobile', 'mobile number', 'alternate mobile number', 'whatsapp number',
    'email', 'official email', 'alternate email', 'e-mail id',
    'rh name', 'relationship owner', 'relationship manager', 'reports to', 'rank', 'team',
    'remarks', 'remark', 'notes', 'eta', 'comment', 'comments',
    'created by', 'created on', 'last modified by', 'last modified on',
    'updated at', 'updated', 'last login', 'password hash', 'password salt',
    'total shipments', 'credit limit', 'credit utilisation', 'designation', 'department',
    'role', 'primary/secondary', 'preferred communication channel', 'states',
    'esc l1', 'esc l2', 'linked seller', 'linked buyer', 'nbfc'
  ],

  LOG_TAB: 'Doc Sync Log',
  MAX_LOG_ROWS: 500,        // log tab is trimmed to this many runs
  MAX_LIST_IN_LOG: 25,      // cap on itemised GSTINs/columns in a single log cell

  AUTO_CREATE_COLUMNS: true, // append a destination column when a doc header is missing
  COPY_BLANKS: true,         // blank in source blanks the destination (verbatim copy)
  CONFLICT: 'first',         // 'first' | 'last' — winner when a GSTIN spans source tabs
  TRIGGER_EVERY_HOURS: 1
};

/* ═══════════════════════════════ Helpers ═══════════════════════════════════ */
/* All helpers are ds*-prefixed to guarantee no collision with the globals
   already defined in Code.gs and MetabaseSync.gs (normKey_, isHiddenDoc_, …). */

/** Normalise a header for comparison: lowercase, collapse punctuation/space. */
function dsNormHeader_(h) {
  return String(h == null ? '' : h)
    .replace(/[\u200B-\u200D\uFEFF]/g, '')   // zero-width junk / BOM from copy-paste
    .trim().toLowerCase()
    .replace(/[_\-/\\.]+/g, ' ')
    .replace(/\s+/g, ' ');
}
/** The one and only key normalisation: trim + uppercase. */
function dsNormGstin_(v) { return String(v == null ? '' : v).trim().toUpperCase(); }

/** Locate the key column in a header row. Returns a 0-based index, or -1. */
function dsFindKeyCol_(headers) {
  var norm = headers.map(dsNormHeader_);
  for (var a = 0; a < DOCSYNC_CFG.KEY_ALIASES.length; a++) {
    var i = norm.indexOf(DOCSYNC_CFG.KEY_ALIASES[a]);
    if (i >= 0) return i;
  }
  // Fall back to any header that merely contains "gst" (e.g. "GSTIN (Seller)").
  for (var j = 0; j < norm.length; j++) if (norm[j].indexOf('gst') !== -1) return j;
  return -1;
}

/** True when a header should be treated as a document column. */
function dsIsDocHeader_(header) {
  var n = dsNormHeader_(header);
  if (!n) return false;
  if (DOCSYNC_CFG.NON_DOC_HEADERS.indexOf(n) !== -1) return false;
  if (DOCSYNC_CFG.KEY_ALIASES.indexOf(n) !== -1) return false;
  // Respect the project's existing hidden-document list when it is loaded.
  if (typeof isHiddenDoc_ === 'function' && isHiddenDoc_(n)) return false;
  return true;
}

/** Group ascending 0-based column indices into contiguous runs → [{start,len}]. */
function dsColumnRuns_(cols) {
  var runs = [], sorted = cols.slice().sort(function (a, b) { return a - b; });
  for (var i = 0; i < sorted.length; i++) {
    var start = sorted[i], len = 1;
    while (i + 1 < sorted.length && sorted[i + 1] === sorted[i] + 1) { len++; i++; }
    runs.push({ start: start, len: len });
  }
  return runs;
}

/** Extract a link URL from a RichTextValue, including single-run cells. */
function dsLinkUrl_(rt) {
  if (!rt) return '';
  try {
    var direct = rt.getLinkUrl();
    if (direct) return direct;
    var runs = rt.getRuns() || [];
    for (var i = 0; i < runs.length; i++) { var u = runs[i].getLinkUrl(); if (u) return u; }
  } catch (e) {}
  return '';
}

/* ═════════════════════════════ Source indexing ═════════════════════════════ */

/**
 * Read every source sheet that carries a GSTIN column and build:
 *   index[GSTIN][normalisedHeader] = { value, formula, link, header, sheet }
 * One getValues()/getFormulas()/getRichTextValues() per sheet — no cell reads.
 */
function dsBuildSourceIndex_(srcSs, stats) {
  var index = {};
  srcSs.getSheets().forEach(function (sh) {
    var name = sh.getName();
    var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (lastRow < 2 || lastCol < 1) { stats.skippedSheets.push(name + ' (source: empty)'); return; }

    var values  = sh.getRange(1, 1, lastRow, lastCol).getValues();
    var headers = values[0];
    var keyCol  = dsFindKeyCol_(headers);
    if (keyCol < 0) { stats.skippedSheets.push(name + ' (source: no GSTIN column)'); return; }

    // Formulas and rich text are only needed to copy links/expressions faithfully.
    var formulas = sh.getRange(1, 1, lastRow, lastCol).getFormulas();
    var rich     = null;
    try { rich = sh.getRange(1, 1, lastRow, lastCol).getRichTextValues(); } catch (e) { rich = null; }

    // Which columns on this sheet are candidate documents?
    var docCols = [];
    for (var c = 0; c < headers.length; c++) {
      if (c === keyCol) continue;
      if (dsIsDocHeader_(headers[c])) docCols.push(c);
    }
    if (!docCols.length) { stats.skippedSheets.push(name + ' (source: no document columns)'); return; }

    stats.sourceSheetsScanned++;
    var seenHere = {};

    for (var r = 1; r < values.length; r++) {
      var gst = dsNormGstin_(values[r][keyCol]);
      if (!gst) { stats.blankKeyRows++; continue; }
      stats.sourceRowsRead++;

      if (seenHere[gst]) {
        stats.duplicateGstins.push(gst + ' (' + name + ' rows ' + (seenHere[gst] + 1) + ' & ' + (r + 1) + ')');
        if (DOCSYNC_CFG.CONFLICT === 'first') continue;   // keep the earlier row
      }
      seenHere[gst] = r;

      var bucket = index[gst] || (index[gst] = {});
      for (var d = 0; d < docCols.length; d++) {
        var col = docCols[d], key = dsNormHeader_(headers[col]);
        var cell = {
          value:   values[r][col],
          formula: formulas[r][col] || '',
          link:    rich ? dsLinkUrl_(rich[r][col]) : '',
          header:  String(headers[col]),
          sheet:   name
        };
        var existing = bucket[key];
        if (!existing) { bucket[key] = cell; continue; }

        // Same GSTIN seen on another tab. Keep per CONFLICT, and note real clashes.
        var a = String(existing.value == null ? '' : existing.value).trim();
        var b = String(cell.value == null ? '' : cell.value).trim();
        if (a !== b && a !== '' && b !== '') {
          stats.crossTabConflicts.push(gst + ' · ' + cell.header + ' (' + existing.sheet + '="' + a + '" vs ' + name + '="' + b + '")');
        }
        if (DOCSYNC_CFG.CONFLICT === 'last' || a === '') bucket[key] = cell;
      }
    }
  });
  return index;
}

/* ═══════════════════════════ Destination writing ═══════════════════════════ */

/** Tabs MetabaseSync clears+rewrites, so overlap can be reported. */
function dsManagedTabs_() {
  try {
    if (typeof CFG !== 'undefined' && CFG && CFG.QUERIES && String(CFG.SHEET_ID) === DOCSYNC_CFG.DEST_ID) {
      return CFG.QUERIES.map(function (q) { return String(q.tab); });
    }
  } catch (e) {}
  return [];
}

/**
 * Apply the source index to one destination sheet.
 * Reads the sheet once, mutates only matched document cells in memory, then
 * writes one setValues() per contiguous run of target columns.
 */
function dsSyncSheet_(sh, index, stats, dryRun) {
  var name = sh.getName();
  if (name === DOCSYNC_CFG.LOG_TAB) return;

  var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastRow < 2 || lastCol < 1) { stats.skippedSheets.push(name + ' (dest: empty)'); return; }

  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var keyCol  = dsFindKeyCol_(headers);
  if (keyCol < 0) { stats.skippedSheets.push(name + ' (dest: no GSTIN column)'); return; }

  /* Every document header present anywhere in the source, so missing
     destination columns can be created on demand. */
  var wanted = {};
  Object.keys(index).forEach(function (g) {
    Object.keys(index[g]).forEach(function (k) { if (!wanted[k]) wanted[k] = index[g][k].header; });
  });

  // Map normalised destination headers → column index.
  var destByNorm = {};
  for (var c = 0; c < headers.length; c++) {
    if (c === keyCol) continue;
    var n = dsNormHeader_(headers[c]);
    if (n && destByNorm[n] == null) destByNorm[n] = c;
  }

  // Targets = document headers that exist in both sheets.
  var targets = [];   // [{ norm, col, header }]
  Object.keys(wanted).forEach(function (n) {
    if (destByNorm[n] != null && dsIsDocHeader_(headers[destByNorm[n]])) {
      targets.push({ norm: n, col: destByNorm[n], header: String(headers[destByNorm[n]]) });
    }
  });

  // Optionally create the columns that are missing, appended to the right.
  var created = [];
  if (DOCSYNC_CFG.AUTO_CREATE_COLUMNS) {
    var missing = Object.keys(wanted).filter(function (n) { return destByNorm[n] == null; });
    if (missing.length && !dryRun) {
      var startCol = sh.getLastColumn() + 1;
      sh.insertColumnsAfter(sh.getLastColumn(), missing.length);
      sh.getRange(1, startCol, 1, missing.length)
        .setValues([missing.map(function (n) { return wanted[n]; })]);
      missing.forEach(function (n, i) {
        targets.push({ norm: n, col: startCol - 1 + i, header: wanted[n] });
        created.push(wanted[n]);
      });
      lastCol = sh.getLastColumn();
    } else if (missing.length && dryRun) {
      missing.forEach(function (n) { created.push(wanted[n] + ' (would create)'); });
    }
  } else {
    Object.keys(wanted).forEach(function (n) {
      if (destByNorm[n] == null) stats.skippedColumns.push(name + ' · ' + wanted[n] + ' (absent in destination)');
    });
  }

  if (!targets.length) { stats.skippedSheets.push(name + ' (dest: no matching document columns)'); return; }
  stats.destSheetsProcessed++;
  targets.forEach(function (t) { stats.docColumns[t.header] = true; });
  created.forEach(function (c) { stats.columnsCreated.push(name + ' · ' + c); });

  // Warn when targets sit inside the block MetabaseSync clears every minute.
  if (dsManagedTabs_().indexOf(name) !== -1) {
    var clobbered = targets.filter(function (t) { return t.col < headers.length; })
                           .map(function (t) { return t.header; });
    if (clobbered.length) {
      stats.warnings.push('"' + name + '": ' + clobbered.length +
        ' column(s) sit inside the range MetabaseSync clears every minute and will be overwritten — ' +
        clobbered.slice(0, DOCSYNC_CFG.MAX_LIST_IN_LOG).join(', '));
    }
  }

  // ── Single batched read of the data region ───────────────────────────────
  var numRows  = lastRow - 1;
  var range    = sh.getRange(2, 1, numRows, lastCol);
  var values   = range.getValues();
  var formulas = range.getFormulas();

  var matched = 0, unmatched = 0, updated = 0;
  var pendingLinks = [];     // [{row, col, text, url}] applied after the value write

  for (var r = 0; r < numRows; r++) {
    var gst = dsNormGstin_(values[r][keyCol]);
    if (!gst) continue;
    var rec = index[gst];
    if (!rec) {
      unmatched++;
      if (stats.unmatchedGstins.length < DOCSYNC_CFG.MAX_LIST_IN_LOG * 4) stats.unmatchedGstins.push(gst + ' (' + name + ')');
      continue;
    }
    matched++;

    for (var t = 0; t < targets.length; t++) {
      var tg = targets[t], src = rec[tg.norm];
      if (!src) continue;                                   // source has no such column
      if (formulas[r][tg.col]) { stats.formulaCellsPreserved++; continue; }   // never clobber a formula

      var next = src.formula ? src.formula : src.value;     // formulas carry =HYPERLINK across
      var isBlank = (next === '' || next == null);
      if (isBlank && !DOCSYNC_CFG.COPY_BLANKS) continue;

      var prev = values[r][tg.col];
      if (String(prev == null ? '' : prev) === String(next == null ? '' : next)) continue;  // already correct

      values[r][tg.col] = next == null ? '' : next;
      updated++;
      if (src.link) pendingLinks.push({ row: r, col: tg.col, text: String(src.value == null ? '' : src.value), url: src.link });
    }
  }

  stats.matched   += matched;
  stats.unmatched += unmatched;

  // ── Batched write-back: one setValues() per contiguous column run ────────
  if (updated && !dryRun) {
    dsColumnRuns_(targets.map(function (t) { return t.col; })).forEach(function (run) {
      var block = values.map(function (row) { return row.slice(run.start, run.start + run.len); });
      sh.getRange(2, run.start + 1, numRows, run.len).setValues(block);
    });

    /* Re-apply rich-text links. Grouped per column so this stays a handful of
       range writes; only columns that actually carry a link are touched, and a
       column is skipped unless all of its written values are text (setRichText
       would otherwise coerce numbers and dates to strings). */
    if (pendingLinks.length) {
      var byCol = {};
      pendingLinks.forEach(function (p) { (byCol[p.col] = byCol[p.col] || []).push(p); });
      Object.keys(byCol).forEach(function (colKey) {
        var col = Number(colKey);
        var allText = true;
        for (var i = 0; i < numRows; i++) {
          var v = values[i][col];
          if (v !== '' && v != null && typeof v !== 'string') { allText = false; break; }
        }
        if (!allText) { stats.warnings.push('Links not applied to "' + headers[col] + '" (column holds non-text values).'); return; }
        var linkAt = {};
        byCol[col].forEach(function (p) { linkAt[p.row] = p; });
        var rt = [];
        for (var r2 = 0; r2 < numRows; r2++) {
          var text = String(values[r2][col] == null ? '' : values[r2][col]);
          var b = SpreadsheetApp.newRichTextValue().setText(text);
          if (linkAt[r2] && text) b.setLinkUrl(linkAt[r2].url);
          rt.push([b.build()]);
        }
        sh.getRange(2, col + 1, numRows, 1).setRichTextValues(rt);
        stats.linksApplied += byCol[col].length;
      });
    }
  }
  stats.cellsUpdated += updated;
}

/* ═══════════════════════════════ Logging ═══════════════════════════════════ */

function dsCap_(arr) {
  if (!arr.length) return '';
  var head = arr.slice(0, DOCSYNC_CFG.MAX_LIST_IN_LOG).join('; ');
  return arr.length > DOCSYNC_CFG.MAX_LIST_IN_LOG
    ? head + ' … (+' + (arr.length - DOCSYNC_CFG.MAX_LIST_IN_LOG) + ' more)' : head;
}

var DS_LOG_HEADERS = ['Timestamp', 'Status', 'Duration (s)', 'Source Sheets', 'Source Rows',
  'Unique GSTINs', 'Duplicate GSTINs', 'Dest Sheets', 'Matched', 'Unmatched',
  'Doc Columns Synced', 'Columns Created', 'Cells Updated', 'Links Applied',
  'Formulas Preserved', 'Skipped', 'Warnings', 'Errors'];

/** Append one row per run to the destination log tab, newest last, trimmed. */
function dsWriteLog_(destSs, s) {
  try {
    var sh = destSs.getSheetByName(DOCSYNC_CFG.LOG_TAB);
    if (!sh) {
      sh = destSs.insertSheet(DOCSYNC_CFG.LOG_TAB);
      sh.getRange(1, 1, 1, DS_LOG_HEADERS.length).setValues([DS_LOG_HEADERS]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    var docCols = Object.keys(s.docColumns);
    sh.appendRow([
      new Date(),
      s.errors.length ? 'ERROR' : (s.warnings.length ? 'OK (warnings)' : 'OK'),
      Math.round(s.durationMs / 100) / 10,
      s.sourceSheetsScanned, s.sourceRowsRead, s.uniqueGstins,
      s.duplicateGstins.length, s.destSheetsProcessed, s.matched, s.unmatched,
      docCols.length + (docCols.length ? ' — ' + dsCap_(docCols) : ''),
      s.columnsCreated.length + (s.columnsCreated.length ? ' — ' + dsCap_(s.columnsCreated) : ''),
      s.cellsUpdated, s.linksApplied, s.formulaCellsPreserved,
      dsCap_(s.skippedSheets.concat(s.skippedColumns)),
      dsCap_(s.warnings.concat(s.crossTabConflicts.length ? ['Cross-tab conflicts: ' + dsCap_(s.crossTabConflicts)] : [])),
      dsCap_(s.errors)
    ]);
    var extra = sh.getLastRow() - 1 - DOCSYNC_CFG.MAX_LOG_ROWS;
    if (extra > 0) sh.deleteRows(2, extra);        // trim oldest runs
  } catch (e) {
    Logger.log('DocSync: could not write the log tab — ' + (e && e.message || e));
  }
}

/** Human-readable summary for the Apps Script execution log. */
function dsLogSummary_(s) {
  Logger.log('──────── DocSync ────────');
  Logger.log('Status              : ' + (s.errors.length ? 'ERROR' : (s.warnings.length ? 'OK (with warnings)' : 'OK')));
  Logger.log('Duration            : ' + (Math.round(s.durationMs / 100) / 10) + 's');
  Logger.log('Source sheets/rows  : ' + s.sourceSheetsScanned + ' / ' + s.sourceRowsRead);
  Logger.log('Unique GSTINs       : ' + s.uniqueGstins);
  Logger.log('Duplicate GSTINs    : ' + s.duplicateGstins.length + (s.duplicateGstins.length ? ' → ' + dsCap_(s.duplicateGstins) : ''));
  Logger.log('Blank-key rows      : ' + s.blankKeyRows);
  Logger.log('Dest sheets         : ' + s.destSheetsProcessed);
  Logger.log('Matched / unmatched : ' + s.matched + ' / ' + s.unmatched);
  Logger.log('Doc columns synced  : ' + Object.keys(s.docColumns).length + (Object.keys(s.docColumns).length ? ' → ' + dsCap_(Object.keys(s.docColumns)) : ''));
  Logger.log('Columns created     : ' + s.columnsCreated.length + (s.columnsCreated.length ? ' → ' + dsCap_(s.columnsCreated) : ''));
  Logger.log('Cells updated       : ' + s.cellsUpdated);
  Logger.log('Links applied       : ' + s.linksApplied);
  Logger.log('Formulas preserved  : ' + s.formulaCellsPreserved);
  if (s.skippedSheets.length)     Logger.log('Skipped sheets      : ' + dsCap_(s.skippedSheets));
  if (s.skippedColumns.length)    Logger.log('Skipped columns     : ' + dsCap_(s.skippedColumns));
  if (s.crossTabConflicts.length) Logger.log('Cross-tab conflicts : ' + dsCap_(s.crossTabConflicts));
  s.warnings.forEach(function (w) { Logger.log('WARNING             : ' + w); });
  s.errors.forEach(function (e) { Logger.log('ERROR               : ' + e); });
  Logger.log('─────────────────────────');
}

/* ═════════════════════════════ Entry points ════════════════════════════════ */

function dsNewStats_() {
  return { sourceSheetsScanned:0, sourceRowsRead:0, uniqueGstins:0, blankKeyRows:0,
           destSheetsProcessed:0, matched:0, unmatched:0, cellsUpdated:0,
           linksApplied:0, formulaCellsPreserved:0,
           duplicateGstins:[], unmatchedGstins:[], crossTabConflicts:[],
           docColumns:{}, columnsCreated:[], skippedSheets:[], skippedColumns:[],
           warnings:[], errors:[], durationMs:0 };
}

/**
 * Synchronise document columns from source to destination.
 * A script lock serialises runs so an overlapping trigger cannot double-write.
 * @param {boolean} dryRun  when true, compute and log the plan but write nothing.
 * @return {Object} the run statistics.
 */
function syncDocuments(dryRun) {
  var t0 = Date.now();
  var stats = dsNewStats_();
  var lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    stats.errors.push('Another DocSync run is in progress — skipped.');
    stats.durationMs = Date.now() - t0;
    dsLogSummary_(stats);
    return stats;
  }
  try {
    var srcSs  = SpreadsheetApp.openById(DOCSYNC_CFG.SOURCE_ID);
    var destSs = SpreadsheetApp.openById(DOCSYNC_CFG.DEST_ID);

    var index = dsBuildSourceIndex_(srcSs, stats);
    stats.uniqueGstins = Object.keys(index).length;
    if (!stats.uniqueGstins) stats.warnings.push('No GSTIN records found in the source — nothing to sync.');

    destSs.getSheets().forEach(function (sh) {
      try {
        dsSyncSheet_(sh, index, stats, !!dryRun);
      } catch (e) {
        // One bad sheet must never abort the rest of the run.
        stats.errors.push('Sheet "' + sh.getName() + '": ' + (e && e.message || e));
      }
    });

    stats.durationMs = Date.now() - t0;
    dsLogSummary_(stats);
    if (!dryRun) dsWriteLog_(destSs, stats);
    return stats;
  } catch (e) {
    stats.errors.push(String(e && e.message || e));
    stats.durationMs = Date.now() - t0;
    dsLogSummary_(stats);
    return stats;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/** Dry run — reports exactly what would change without touching the sheet. */
function previewDocSyncPlan() { return syncDocuments(true); }

/** Correct ordering when MetabaseSync owns the same tabs: refresh, then sync. */
function syncMetabaseThenDocs() {
  if (typeof syncMetabaseToSheet === 'function') {
    try { syncMetabaseToSheet(); }
    catch (e) { Logger.log('DocSync: Metabase refresh failed, continuing — ' + (e && e.message || e)); }
  }
  return syncDocuments();
}

/** Install the recurring trigger. Idempotent — clears its own triggers first. */
function installDocSyncTrigger() {
  removeDocSyncTriggers();
  ScriptApp.newTrigger('syncDocuments')
    .timeBased()
    .everyHours(DOCSYNC_CFG.TRIGGER_EVERY_HOURS)
    .create();
  Logger.log('DocSync: trigger installed — syncDocuments() every ' +
             DOCSYNC_CFG.TRIGGER_EVERY_HOURS + 'h.');
}

/** Remove every trigger this script owns. */
function removeDocSyncTriggers() {
  var n = 0;
  ScriptApp.getProjectTriggers().forEach(function (t) {
    var fn = t.getHandlerFunction();
    if (fn === 'syncDocuments' || fn === 'syncMetabaseThenDocs') { ScriptApp.deleteTrigger(t); n++; }
  });
  if (n) Logger.log('DocSync: removed ' + n + ' trigger(s).');
}
