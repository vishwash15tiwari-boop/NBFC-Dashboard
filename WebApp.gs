// ============================================================
//  NBFC Dashboard — Web App backend
//
//  Serves Index.html as a live web dashboard and returns the
//  computed document-collection data as JSON. Reuses the helpers
//  and constants declared in Code.gs (same Apps Script project,
//  shared global scope): book_, locateSource_, RECEIVED_VALUES,
//  NA_VALUES.
//
//  DEPLOY:
//    Apps Script editor -> Deploy -> New deployment
//      Type: Web app
//      Execute as: Me
//      Who has access: (your choice — see README)
//    -> copy the Web app URL.
// ============================================================

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('NBFC Document Dashboard')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** Allows Index.html to pull in extra HTML partials if ever split out. */
function include(name) {
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/** Called from the frontend via google.script.run. */
function getDashboardData() {
  try {
    return computeDashboard_();
  } catch (err) {
    return { error: String(err && err.message ? err.message : err) };
  }
}

// ── read the sheet and compute everything (values, not formulas) ──

function computeDashboard_() {
  const ss  = book_();
  const src = locateSource_(ss);
  const sh  = src.sheet;

  const recvSet = RECEIVED_VALUES.map(s => s.toLowerCase());
  const naSet   = NA_VALUES.map(s => s.toLowerCase());
  const classify = v => {
    const t = String(v == null ? '' : v).trim().toLowerCase();
    if (recvSet.indexOf(t) > -1) return 'recv';
    if (naSet.indexOf(t) > -1)   return 'na';
    return 'pend'; // No / "-" / blank / anything else
  };

  const N        = src.numDocs;
  const lastRow  = sh.getLastRow();
  const numRows  = Math.max(lastRow - src.dataStart + 1, 0);

  const base = {
    sourceName: src.name,
    generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm'),
    numDocs: N,
    docLabels: src.docLabels,
    sellers: [],
    docs: src.docLabels.map(l => ({ label: l, recv: 0, na: 0, appl: 0, pend: 0, pct: 0 })),
    totals: { sellers: 0, recv: 0, na: 0, appl: 0, pend: 0, pct: 0 },
  };
  if (numRows === 0) return base;

  const names    = sh.getRange(src.dataStart, src.nameCol, numRows, 1).getValues();
  const entities = src.entityCol
    ? sh.getRange(src.dataStart, src.entityCol, numRows, 1).getValues()
    : null;
  const block = sh.getRange(src.dataStart, src.docStart, numRows, N).getValues();

  let tRecv = 0, tNa = 0;

  for (let i = 0; i < numRows; i++) {
    const name = String(names[i][0] == null ? '' : names[i][0]).trim();
    if (!name) continue; // skip blank / buyer-only rows

    let recv = 0, na = 0;
    for (let j = 0; j < N; j++) {
      const cls = classify(block[i][j]);
      if (cls === 'recv') { recv++; base.docs[j].recv++; }
      else if (cls === 'na') { na++; base.docs[j].na++; }
    }
    const appl = N - na;
    const pend = appl - recv;
    base.sellers.push({
      name: name,
      type: entities ? String(entities[i][0] || '').trim() : '',
      recv: recv, na: na, appl: appl, pend: pend,
      pct: appl ? Math.round((recv / appl) * 100) : 0,
    });
    tRecv += recv;
    tNa   += na;
  }

  const S = base.sellers.length;
  base.docs.forEach(d => {
    d.appl = S - d.na;
    d.pend = d.appl - d.recv;
    d.pct  = d.appl ? Math.round((d.recv / d.appl) * 100) : 0;
  });

  const tAppl = N * S - tNa;
  base.totals = {
    sellers: S,
    recv: tRecv,
    na: tNa,
    appl: tAppl,
    pend: tAppl - tRecv,
    pct: tAppl ? Math.round((tRecv / tAppl) * 1000) / 10 : 0, // one decimal
  };
  return base;
}
