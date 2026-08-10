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

/* ── Seller / Buyer NBFC classification ─────────────────────────────────────
   Single source of truth for which NBFCs belong to each entity type.
   To add a new NBFC to a group, update this map and CONFIG.NBFC_TABS only —
   no other logic changes are required.                                        */
var NBFC_ENTITY_MAP = {
  seller: ['billmart', 'capitalxb', 'karncy'],
  buyer:  ['strideone', 'credable']
};

var CONFIG = {
  // Native Google Sheet ID (already confirmed native; no xlsx conversion needed).
  SOURCE_FILE_ID: '1RoHWbZyHhNKlweWXD4AMSZfB5ONdktPcVayOkpPgjpo',

  // One entry per NBFC tracker tab.
  // maxCol caps how far right the tab is read so scratch columns to the right
  // never pollute KPIs, cards, the matrix, or form saves.
  // docStartCol / docEndCol (1-indexed) restrict which columns are treated as
  // document-status columns — L=12 through AA=27.
  NBFC_TABS: [
    { id: 'billmart',  name: 'Billmart',   maxCol: 40, docStartCol: 12, docEndCol: 27, entityGroup: 'seller' },
    { id: 'capitalxb', name: 'Capital XB', maxCol: 40, docStartCol: 12, docEndCol: 27, entityGroup: 'seller' },
    { id: 'strideone', name: 'StrideOne',  maxCol: 40, docStartCol: 12, docEndCol: 27, entityGroup: 'buyer' },
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

/* ═══════════════════════════════════════════════════════════════════════════
   NBFC PLATFORM  ·  live data bridge for Index.html
   ───────────────────────────────────────────────────────────────────────────
   Read  : getPlatformData()  → sellers + buyers (from the Metabase-synced
           Sellers/Buyers tabs) plus any document statuses ops have entered.
   Write : savePlatformDoc()  → upserts one document status.

   Identity (name/GSTIN/region/…) is owned by Metabase, which rewrites the
   Sellers/Buyers tabs every minute (see MetabaseSync.gs) — those tabs are
   read-only from the dashboard's side. Anything ops ENTER in the dashboard is
   stored in a SEPARATE tab (PLATFORM_DOCSTATUS_TAB) so the next Metabase sync
   can never overwrite it. Read joins the two back together by GSTIN.
   ═══════════════════════════════════════════════════════════════════════════ */

var PLATFORM_SHEET_ID      = MB_SHEET_ID;   // 1d57KGl00-… — the Metabase-synced master sheet
var PLATFORM_DOCSTATUS_TAB = 'DocStatus';

/** Reads one identity tab ('Sellers' | 'Buyers') into plain records. */
function readPlatformTab_(ss, tabName, type) {
  var sh = ss.getSheetByName(tabName);
  if (!sh) return [];
  var vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];

  var idx = {};
  vals[0].forEach(function (h, i) { idx[String(h).trim().toLowerCase()] = i; });
  function col(row, names) {
    for (var i = 0; i < names.length; i++) {
      var k = names[i].toLowerCase();
      if (k in idx) {
        var v = row[idx[k]];
        if (v != null && String(v).trim() !== '') return String(v).trim();
      }
    }
    return '';
  }
  var nameCols = type === 'seller'
    ? ['Seller Business Name', 'Business Name', 'Entity Name']
    : ['Buyer Business Name', 'Business Name', 'Entity Name'];
  var typeCols = type === 'seller' ? ['Seller Type'] : ['Buyer Type', 'Seller Type'];

  var out = [];
  for (var r = 1; r < vals.length; r++) {
    var row  = vals[r];
    var gst  = col(row, ['GSTIN', 'GST', 'Seller GSTIN', 'Buyer GSTIN']);
    var name = col(row, nameCols);
    if (!gst && !name) continue;                 // skip blank rows
    out.push({
      type:         type,
      name:         name,
      gstin:        gst,
      region:       col(row, ['Region']),
      vertical:     col(row, ['Vertical']),
      businessType: col(row, typeCols),
      state:        col(row, ['State']),
      vintage:      col(row, ['Vintage with Recykal', 'Vintage']),
      entityType:   col(row, ['Entity Type']),
      onboarding:   col(row, ['Onboarding Status', 'Status']) || 'Completed'
    });
  }
  return out;
}

/** Reads the DocStatus tab into { "<GSTIN>||<docId>": { status, comment } }. */
function readPlatformDocStatus_(ss) {
  var map = {};
  var sh = ss.getSheetByName(PLATFORM_DOCSTATUS_TAB);
  if (!sh) return map;
  var vals = sh.getDataRange().getValues();
  if (vals.length < 2) return map;
  var H  = vals[0].map(function (h) { return String(h).trim().toLowerCase(); });
  var gi = H.indexOf('gstin'), di = H.indexOf('doc id'), si = H.indexOf('status'), ci = H.indexOf('comment');
  if (gi < 0 || di < 0 || si < 0) return map;
  for (var r = 1; r < vals.length; r++) {
    var g = String(vals[r][gi]).trim().toUpperCase();
    var d = String(vals[r][di]).trim();
    if (g && d) map[g + '||' + d] = {
      status:  String(vals[r][si]).trim(),
      comment: ci >= 0 ? String(vals[r][ci] == null ? '' : vals[r][ci]).trim() : ''
    };
  }
  return map;
}

/**
 * Called from Index.html on load. Returns sellers + buyers (identity from the
 * Metabase-synced tabs) plus the document statuses ops have entered.
 */
function getPlatformData(token) {
  try {
    var me = requireSession_(token);
    var ss = SpreadsheetApp.openById(PLATFORM_SHEET_ID);

    var sellers = applyScope_(me, readPlatformTab_(ss, 'Sellers', 'seller'));
    var buyers  = applyScope_(me, readPlatformTab_(ss, 'Buyers',  'buyer'));

    /* Document statuses are keyed by GSTIN, so they need the same narrowing —
       otherwise an out-of-scope vendor's document history would still ship. */
    var visible = {};
    sellers.concat(buyers).forEach(function (r) { visible[String(r.gstin || '').toUpperCase()] = true; });
    var all = readPlatformDocStatus_(ss), docStatus = {};
    Object.keys(all).forEach(function (k) { if (visible[k.split('||')[0]]) docStatus[k] = all[k]; });

    return {
      ok:          true,
      sellers:     sellers,
      buyers:      buyers,
      docStatus:   docStatus,
      profile:     scrubClientPayload_(me),
      scoped:      !me.perms.viewAllStates,
      generatedAt: new Date().toISOString()
    };
  } catch (e) {
    var msg = String(e && e.message || e);
    if (msg === 'AUTH_REQUIRED') {
      return { ok:false, authRequired:true, error:'Your session has expired. Please sign in again.',
               sellers:[], buyers:[], docStatus:{} };
    }
    return { ok:false, error:msg, sellers:[], buyers:[], docStatus:{} };
  }
}

/**
 * Called from Index.html when ops set a document's status or comment. Upserts
 * one row in the DocStatus tab, keyed by (GSTIN, docId). Creates the tab on
 * first write. payload = { gstin, entityType, docId, docName, status, comment }
 */
function savePlatformDoc(payload, token) {
  try {
    var me = requireSession_(token);
    if (!me.perms.editDocs) return { ok:false, error:'Your role cannot edit documents.' };
    payload = payload || {};
    var gst   = String(payload.gstin || '').trim().toUpperCase();
    var docId = String(payload.docId || '').trim();
    if (!gst)   return { ok: false, error: 'GSTIN is required' };
    if (!docId) return { ok: false, error: 'docId is required' };

    var ss = SpreadsheetApp.openById(PLATFORM_SHEET_ID);
    var sh = ss.getSheetByName(PLATFORM_DOCSTATUS_TAB);
    if (!sh) {
      sh = ss.insertSheet(PLATFORM_DOCSTATUS_TAB);
      sh.getRange(1, 1, 1, 7).setValues([['GSTIN', 'Entity Type', 'Doc ID', 'Doc Name', 'Status', 'Comment', 'Updated At']]);
      sh.setFrozenRows(1);
    }

    var vals   = sh.getDataRange().getValues();
    var rowNum = -1;                              // find existing (GSTIN, docId)
    for (var r = 1; r < vals.length; r++) {
      if (String(vals[r][0]).trim().toUpperCase() === gst &&
          String(vals[r][2]).trim() === docId) { rowNum = r + 1; break; }
    }
    var rowVals = [gst, payload.entityType || '', docId, payload.docName || '',
                   payload.status || '', payload.comment || '', new Date()];
    if (rowNum > 0) sh.getRange(rowNum, 1, 1, 7).setValues([rowVals]);
    else            sh.appendRow(rowVals);

    return { ok: true };
  } catch (e) {
    return { ok: false, error: String(e && e.message || e) };
  }
}

/* ═══════════════════════════════════════════════════════════════════════════
   POC DIRECTORY  ·  live storage for Index.html's POC module (additive)
   ───────────────────────────────────────────────────────────────────────────
   Two tabs, both untouched by the Metabase sync:
     PLATFORM_POC_TAB    — one row per contact, keyed by POC ID
     PLATFORM_POCFU_TAB  — one row per logged interaction
   getPOCData() joins them; savePOC() upserts a contact; savePOCFollowup()
   appends an interaction. Nothing here reads or writes the existing tabs.
   ═══════════════════════════════════════════════════════════════════════════ */
var PLATFORM_POC_TAB   = 'POC Directory';
var PLATFORM_POCFU_TAB = 'POC Followups';

/* ── Canonical POC Directory schema ──────────────────────────────────────────
   Single source of truth for the sheet's column layout, the read mapping, AND
   the write mapping. `header` order defines the columns of a freshly-created
   tab; `aliases` (lower-cased) let an EXISTING tab with older or differently
   named headers map cleanly — strict one-to-one, so no column is ever ignored.
   Add a field here and it flows to the sheet, read, write, and audit with no
   other backend change. */
var POC_FIELDS = [
  { key:'id',            header:'POC ID',                          aliases:['poc id','id'] },
  { key:'nbfc',          header:'NBFC',                            aliases:['nbfc'] },
  { key:'entityType',    header:'Entity Type',                     aliases:['entity type','seller/buyer','entitytype','type','entity'] },
  { key:'linkedSeller',  header:'Linked Seller',                   aliases:['linked seller','seller','seller name','linked seller name'] },
  { key:'linkedBuyer',   header:'Linked Buyer',                    aliases:['linked buyer','buyer','buyer name','linked buyer name'] },
  { key:'org',           header:'Company Name',                    aliases:['company name','company','organization','org'] },
  { key:'name',          header:'POC Name',                        aliases:['poc name','name','contact name'] },
  { key:'designation',   header:'Designation',                     aliases:['designation'] },
  { key:'department',    header:'Department',                      aliases:['department'] },
  { key:'mobile',        header:'Mobile Number',                   aliases:['mobile number','mobile'] },
  { key:'altMobile',     header:'Alternate Mobile Number',         aliases:['alternate mobile number','alt mobile','alternate mobile','alt mobile number'] },
  { key:'email',         header:'Official Email',                  aliases:['official email','email','email address'] },
  { key:'altEmail',      header:'Alternate Email',                 aliases:['alternate email','alt email','alternate email address'] },
  { key:'whatsapp',      header:'WhatsApp Number',                 aliases:['whatsapp number','whatsapp','wa number','whatsapp no'] },
  { key:'channel',       header:'Preferred Communication Channel', aliases:['preferred communication channel','preferred channel','channel'] },
  { key:'primary',       header:'Primary/Secondary',               aliases:['primary/secondary','primary','primary poc'] },
  { key:'role',          header:'Role',                            aliases:['role','role/purpose','purpose'] },
  { key:'relOwner',      header:'Relationship Owner',              aliases:['relationship owner','relationship manager','rel owner'] },
  { key:'status',        header:'Status',                          aliases:['status'] },
  { key:'lastContact',   header:'Last Interaction Date',           aliases:['last interaction date','last contact','last interaction','last contact date'] },
  { key:'nextFollowup',  header:'Next Follow-up Date',             aliases:['next follow-up date','next followup','next follow-up','next follow up','next followup date'] },
  { key:'followupStatus',header:'Follow-up Status',                aliases:['follow-up status','followup status','fu status'] },
  { key:'remarks',       header:'Remarks',                         aliases:['remarks','remarks/notes','notes'] },
  { key:'createdBy',     header:'Created By',                       aliases:['created by'] },
  { key:'createdOn',     header:'Created On',                      aliases:['created on','created date','created at'] },
  { key:'modifiedBy',    header:'Last Modified By',                aliases:['last modified by','modified by','updated by'] },
  { key:'modifiedOn',    header:'Last Modified On',                aliases:['last modified on','modified on','updated at','updated','last modified'] },

  /* ── Access management ────────────────────────────────────────────────────
     Drives who may log in and what they may see. Safe to send to the client:
     these are directory attributes, not credentials. `active` is the login
     gate; `states` scopes an Operations user's visible rows; `reportsTo`/`rank`
     build the hierarchy; `escL1`/`escL2` are read live so escalation addresses
     are never hardcoded anywhere in the app. */
  { key:'team',          header:'Team',                            aliases:['team','department team','business team'] },
  { key:'states',        header:'States',                          aliases:['states','state','assigned states','territory'] },
  { key:'reportsTo',     header:'Reports To',                      aliases:['reports to','manager','reporting manager','reports to email'] },
  { key:'rank',          header:'Rank',                            aliases:['rank','level','seniority'] },
  { key:'escL1',         header:'Esc L1',                          aliases:['esc l1','escalation level 1','escalation 1','escalation l1'] },
  { key:'escL2',         header:'Esc L2',                          aliases:['esc l2','escalation level 2','escalation 2','escalation l2'] },
  { key:'active',        header:'Active',                          aliases:['active','active status','is active','enabled'] }
];
var POC_HEADERS   = POC_FIELDS.map(function (f) { return f.header; });

/* ── Credential columns ──────────────────────────────────────────────────────
   Deliberately NOT part of POC_FIELDS. getPOCData() builds its client payload
   by walking POC_FIELDS, so keeping these out of that list is what guarantees a
   hash can never be serialized to the browser by accident. They live in the same
   POC Directory tab (one sheet, as specified) but are read only by the auth code
   below, and scrubClientPayload_() is a second, independent guard. */
var POC_AUTH_FIELDS = [
  { key:'pwHash',    header:'Password Hash', aliases:['password hash','pw hash','passwordhash'] },
  { key:'pwSalt',    header:'Password Salt', aliases:['password salt','pw salt','passwordsalt'] },
  { key:'lastLogin', header:'Last Login',    aliases:['last login','last login at','last signed in'] }
];
var POC_AUTH_HEADERS = POC_AUTH_FIELDS.map(function (f) { return f.header; });

/** Defence in depth: strip anything credential-shaped from a client payload,
    whatever list it arrived through. Runs on every authenticated response. */
function scrubClientPayload_(obj) {
  if (obj == null || typeof obj !== 'object') return obj;
  if (Object.prototype.toString.call(obj) === '[object Array]') return obj.map(scrubClientPayload_);
  var out = {};
  Object.keys(obj).forEach(function (k) {
    if (/pass|hash|salt|secret|token|pepper|credential/i.test(k)) return;
    out[k] = scrubClientPayload_(obj[k]);
  });
  return out;
}
var POCFU_HEADERS = ['Followup ID','POC ID','NBFC','Date','Mode','Summary','Next Action','Next Date','By','Created At'];

/** Map the sheet's current headers → { fieldKey: colIndex } via the alias table. */
function pocColIndex_(headers) {
  var H = {}; headers.forEach(function (h, i) { H[String(h).trim().toLowerCase()] = i; });
  var CI = {};
  POC_FIELDS.forEach(function (f) {
    for (var i = 0; i < f.aliases.length; i++) { if (H[f.aliases[i]] != null) { CI[f.key] = H[f.aliases[i]]; break; } }
  });
  return CI;
}

/** Map the sheet's headers → { authFieldKey: colIndex }. Server-side only. */
function pocAuthColIndex_(headers) {
  var H = {}; headers.forEach(function (h, i) { H[String(h).trim().toLowerCase()] = i; });
  var CI = {};
  POC_AUTH_FIELDS.forEach(function (f) {
    for (var i = 0; i < f.aliases.length; i++) { if (H[f.aliases[i]] != null) { CI[f.key] = H[f.aliases[i]]; break; } }
  });
  return CI;
}

/** Guarantee every canonical field has a column, appending any missing ones to
    the right. Never reorders or deletes existing columns, so an existing tab's
    formatting and architecture are preserved. Returns a fresh CI map.
    Credential columns are provisioned alongside so a tab is login-ready, but
    they are indexed separately and never enter the client mapping. */
function ensurePocColumns_(sh) {
  var lastCol = Math.max(1, sh.getLastColumn());
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var CI = pocColIndex_(headers), ACI = pocAuthColIndex_(headers);
  var toAdd = [];
  POC_FIELDS.forEach(function (f)      { if (CI[f.key]  == null) toAdd.push(f.header); });
  POC_AUTH_FIELDS.forEach(function (f) { if (ACI[f.key] == null) toAdd.push(f.header); });
  if (toAdd.length) {
    sh.getRange(1, lastCol + 1, 1, toAdd.length).setValues([toAdd]);
    sh.setFrozenRows(1);
    CI = pocColIndex_(headers.concat(toAdd));
  }
  return CI;
}

/** One field's sheet-cell value from a POC object (normalises Primary/Status). */
function pocFieldValue_(poc, key) {
  if (key === 'primary') return poc.primary ? 'Yes' : 'No';
  if (key === 'status')  return poc.status || 'Active';
  return poc[key] != null ? poc[key] : '';
}
function pocRowFrom_(poc) {
  return POC_FIELDS.map(function (f) { return pocFieldValue_(poc, f.key); });
}

/** Which entity group ('seller'|'buyer') an NBFC id belongs to, per the
    authoritative NBFC_ENTITY_MAP. '' when unknown. */
function nbfcEntityGroup_(nbfcId) {
  var id = String(nbfcId || '').trim().toLowerCase();
  if (!id) return '';
  var group = '';
  Object.keys(NBFC_ENTITY_MAP).forEach(function (g) {
    if (NBFC_ENTITY_MAP[g].map(function (x) { return String(x).trim().toLowerCase(); }).indexOf(id) >= 0) group = g;
  });
  return group;
}

/* ═══════════════════════════════════════════════════════════════════════════
   LOGIN & ACCESS MANAGEMENT
   ───────────────────────────────────────────────────────────────────────────
   Every rule below is enforced HERE, on the server. Nothing in Index.html is
   trusted: the browser can call any google.script.run endpoint directly and can
   be edited at will, so a check that lives in the client is decoration. The
   contract is therefore:
     · the password hash never leaves this file
     · every data endpoint requires a session token and re-derives the caller's
       permissions from the sheet on each call
     · rows outside the caller's scope are removed BEFORE the response is built,
       so restricted data is never delivered and then hidden
   The POC Directory sheet is the only source of users, roles, states and
   escalation contacts — adding, changing or deactivating a person there takes
   effect on their next request with no code change.
   ═══════════════════════════════════════════════════════════════════════════ */

var AUTH_PEPPER_KEY  = 'NBFC_AUTH_PEPPER';
var AUTH_ROUNDS      = 1000;   // iterated SHA-256; ~0.3s in Apps Script
var SESSION_TTL_SEC  = 1800;   // 30-minute sliding inactivity window
var AUTH_MAX_ATTEMPTS = 6;     // per email, per lockout window
var AUTH_LOCKOUT_SEC  = 900;   // 15-minute lockout after repeated failures

/** Server-only secret mixed into every hash, so sheet contents alone are not
    enough to mount an offline attack. Generated once, stored in Script
    Properties (never in the sheet, never in the repo). */
function authPepper_() {
  var props = PropertiesService.getScriptProperties();
  var p = props.getProperty(AUTH_PEPPER_KEY);
  if (!p) { p = Utilities.getUuid() + Utilities.getUuid(); props.setProperty(AUTH_PEPPER_KEY, p); }
  return p;
}

function bytesToHex_(bytes) {
  var s = '';
  for (var i = 0; i < bytes.length; i++) s += ((bytes[i] & 0xFF) + 0x100).toString(16).slice(1);
  return s;
}

/** Salted, peppered, iterated SHA-256. The iteration count is what makes a
    leaked sheet expensive to attack; the per-user salt stops one cracked
    password revealing every identical one. */
function hashPassword_(password, salt) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(salt) + '|' + String(password) + '|' + authPepper_(),
    Utilities.Charset.UTF_8);
  for (var i = 1; i < AUTH_ROUNDS; i++) {
    bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes);
  }
  return bytesToHex_(bytes);
}

/** Comparison whose duration does not depend on where the first difference is,
    so a caller cannot learn the hash byte-by-byte from response timing. */
function safeEqual_(a, b) {
  a = String(a || ''); b = String(b || '');
  if (a.length !== b.length) return false;
  var diff = 0;
  for (var i = 0; i < a.length; i++) diff |= (a.charCodeAt(i) ^ b.charCodeAt(i));
  return diff === 0;
}

function newSalt_() { return Utilities.getUuid().replace(/-/g, ''); }

/* ── Role model ────────────────────────────────────────────────────────────
   This rollout ships two configured tiers, Admin and Operations. Anything the
   sheet does not recognise — blank, "-", a role added later — falls to
   'readonly' rather than being granted Operations by default: an unrecognised
   role must never widen access. */
var ROLE_PERMISSIONS = {
  admin: {
    label:'Admin', viewAllStates:true, viewAllTeams:true,
    editDocs:true, managePOC:true, configure:true, manageUsers:true, exportData:true,
    routes:'*'
  },
  operations: {
    label:'Operations', viewAllStates:false, viewAllTeams:false,
    editDocs:true, managePOC:true, configure:false, manageUsers:false, exportData:true,
    routes:['dashboard','sellers','buyers','matrix','poc','nbfc','profile']
  },
  readonly: {
    label:'Read Only', viewAllStates:false, viewAllTeams:false,
    editDocs:false, managePOC:false, configure:false, manageUsers:false, exportData:false,
    routes:['dashboard','sellers','buyers','matrix','poc','nbfc','profile']
  }
};

function roleKey_(role) {
  var r = String(role == null ? '' : role).trim().toLowerCase();
  if (!r || r === '-') return 'readonly';
  if (r === 'admin' || r === 'administrator' || r === 'super admin') return 'admin';
  if (r === 'operations' || r === 'ops' || r === 'operation' ||
      r === 'poc' || r === 'sub admin' || r === 'subadmin') return 'operations';
  return 'readonly';                       // unknown role → least privilege
}

/** "Yes"/"TRUE"/"1"/"Active" all mean active; blank or anything else does not.
    Blank must read as inactive — the spec requires No/blank to be denied. */
function isActiveFlag_(v) {
  var s = String(v == null ? '' : v).trim().toLowerCase();
  return s === 'yes' || s === 'y' || s === 'true' || s === '1' || s === 'active';
}

/** Split a States cell into normalised names. Accepts comma, semicolon, slash,
    pipe or newline separators so the sheet author is not constrained. */
function splitStates_(v) {
  return String(v == null ? '' : v).split(/[,;/|\n]+/)
    .map(function (s) { return s.trim().toLowerCase(); })
    .filter(function (s) { return s && s !== '-'; });
}

/** Load one directory row by email, including credential columns.
    Server-side only — the return value must never be handed to the client. */
function findDirectoryUser_(email) {
  var key = String(email || '').trim().toLowerCase();
  if (!key) return null;
  var ss = SpreadsheetApp.openById(PLATFORM_SHEET_ID);
  var sh = ss.getSheetByName(PLATFORM_POC_TAB);
  if (!sh) return null;
  var vals = sh.getDataRange().getValues();
  if (vals.length < 2) return null;

  var CI  = pocColIndex_(vals[0]);
  var ACI = pocAuthColIndex_(vals[0]);
  if (CI.email == null) return null;

  for (var r = 1; r < vals.length; r++) {
    if (String(vals[r][CI.email] || '').trim().toLowerCase() !== key) continue;
    function f(i)  { return i == null ? '' : String(vals[r][i] == null ? '' : vals[r][i]).trim(); }
    return {
      rowIndex:    r + 1,
      email:       f(CI.email),
      name:        f(CI.name),
      designation: f(CI.designation),
      role:        f(CI.role),
      team:        f(CI.team),
      states:      f(CI.states),
      reportsTo:   f(CI.reportsTo),
      rank:        f(CI.rank),
      escL1:       f(CI.escL1),
      escL2:       f(CI.escL2),
      activeRaw:   f(CI.active),
      pwHash:      f(ACI.pwHash),
      pwSalt:      f(ACI.pwSalt),
      lastLogin:   f(ACI.lastLogin)
    };
  }
  return null;
}

/** The client-facing shape of a signed-in user. Credentials are absent by
    construction, and the permission block is derived here so the browser never
    decides its own rights. */
function publicProfile_(u) {
  var rk = roleKey_(u.role);
  return {
    email:       u.email,
    name:        u.name || u.email,
    designation: u.designation,
    role:        u.role || '-',
    roleKey:     rk,
    roleLabel:   ROLE_PERMISSIONS[rk].label,
    team:        u.team,
    states:      splitStates_(u.states),
    statesRaw:   u.states,
    reportsTo:   u.reportsTo,
    rank:        u.rank,
    escL1:       u.escL1,
    escL2:       u.escL2,
    active:      isActiveFlag_(u.activeRaw),
    lastLogin:   u.lastLogin,
    perms:       ROLE_PERMISSIONS[rk]
  };
}

/* ── Brute-force throttle ─────────────────────────────────────────────────── */
function attemptKey_(email) { return 'authfail_' + String(email || '').trim().toLowerCase(); }
function bumpFailures_(email) {
  var c = CacheService.getScriptCache(), k = attemptKey_(email);
  var n = parseInt(c.get(k) || '0', 10) + 1;
  c.put(k, String(n), AUTH_LOCKOUT_SEC);
  return n;
}
function isLockedOut_(email) {
  return parseInt(CacheService.getScriptCache().get(attemptKey_(email)) || '0', 10) >= AUTH_MAX_ATTEMPTS;
}
function clearFailures_(email) { CacheService.getScriptCache().remove(attemptKey_(email)); }

/* ── Sessions ─────────────────────────────────────────────────────────────── */
function newSession_(profile) {
  var token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('sess_' + token, JSON.stringify({ email: profile.email }), SESSION_TTL_SEC);
  return token;
}

/** Resolve a token to a live profile. Re-reads the directory every call, so a
    deactivation, role change or state reassignment takes effect immediately
    rather than lasting until the session expires. Touching the cache entry
    gives the sliding inactivity expiry the spec asks for. */
function validateSession_(token) {
  if (!token) return null;
  var c = CacheService.getScriptCache(), k = 'sess_' + String(token);
  var raw = c.get(k);
  if (!raw) return null;
  var email;
  try { email = JSON.parse(raw).email; } catch (e) { return null; }
  var u = findDirectoryUser_(email);
  if (!u || !isActiveFlag_(u.activeRaw)) { c.remove(k); return null; }   // revoked mid-session
  c.put(k, raw, SESSION_TTL_SEC);
  return publicProfile_(u);
}

/** Throws on an invalid session so no endpoint can forget to check. */
function requireSession_(token) {
  var p = validateSession_(token);
  if (!p) throw new Error('AUTH_REQUIRED');
  return p;
}

/**
 * Called from the login screen. Returns { ok, token, profile } or a reason.
 * Failure messages are deliberately identical for "no such user", "wrong
 * password" and "no password set" so the response cannot be used to enumerate
 * who exists in the directory.
 */
function authenticate(email, password) {
  try {
    var key = String(email || '').trim().toLowerCase();
    if (!key || !password) return { ok:false, error:'Enter your email and password.' };
    if (isLockedOut_(key)) {
      return { ok:false, error:'Too many failed attempts. Try again in 15 minutes.' };
    }
    var u = findDirectoryUser_(key);
    var GENERIC = 'Email or password is incorrect.';

    if (!u)                       { bumpFailures_(key); return { ok:false, error:GENERIC }; }
    if (!isActiveFlag_(u.activeRaw)) {
      /* Named explicitly: an active-status denial is not a credential hint, and
         the person needs to know to contact their admin rather than retry. */
      return { ok:false, error:'This account is not active. Contact your administrator.' };
    }
    if (!u.pwHash || !u.pwSalt)   { bumpFailures_(key); return { ok:false, error:GENERIC }; }
    if (!safeEqual_(hashPassword_(password, u.pwSalt), u.pwHash)) {
      bumpFailures_(key); return { ok:false, error:GENERIC };
    }

    clearFailures_(key);
    var profile = publicProfile_(u);
    stampLastLogin_(u);
    profile.lastLogin = u.lastLogin;          // show the PREVIOUS login, not this one
    return { ok:true, token:newSession_(profile), profile:scrubClientPayload_(profile) };
  } catch (e) {
    return { ok:false, error:'Sign-in failed: ' + String(e && e.message || e) };
  }
}

function stampLastLogin_(u) {
  try {
    var ss = SpreadsheetApp.openById(PLATFORM_SHEET_ID);
    var sh = ss.getSheetByName(PLATFORM_POC_TAB);
    if (!sh) return;
    var ACI = pocAuthColIndex_(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
    if (ACI.lastLogin == null) return;
    sh.getRange(u.rowIndex, ACI.lastLogin + 1).setValue(new Date());
  } catch (e) { /* a login must not fail because the stamp could not be written */ }
}

function logout(token) {
  if (token) CacheService.getScriptCache().remove('sess_' + String(token));
  return { ok:true };
}

/** Re-read the caller's own profile (used on resume and by My Profile). */
function getMyProfile(token) {
  try { return { ok:true, profile:scrubClientPayload_(requireSession_(token)) }; }
  catch (e) { return { ok:false, error:String(e && e.message || e) }; }
}

/** A user changing their own password. Requires the current one. */
function changeMyPassword(token, currentPassword, newPassword) {
  try {
    var me = requireSession_(token);
    var u  = findDirectoryUser_(me.email);
    if (!u) return { ok:false, error:'Account not found.' };
    if (!u.pwHash || !safeEqual_(hashPassword_(currentPassword, u.pwSalt), u.pwHash)) {
      return { ok:false, error:'Current password is incorrect.' };
    }
    var err = passwordPolicyError_(newPassword);
    if (err) return { ok:false, error:err };
    writeCredential_(u.rowIndex, newPassword);
    return { ok:true };
  } catch (e) { return { ok:false, error:String(e && e.message || e) }; }
}

function passwordPolicyError_(pw) {
  pw = String(pw || '');
  if (pw.length < 8) return 'Password must be at least 8 characters.';
  if (!/[A-Za-z]/.test(pw) || !/[0-9]/.test(pw)) return 'Password must contain a letter and a number.';
  return '';
}

/** Write a fresh salt+hash for one directory row. The plaintext is used here
    and discarded; it is never stored, logged or returned. */
function writeCredential_(rowIndex, plaintext) {
  var ss = SpreadsheetApp.openById(PLATFORM_SHEET_ID);
  var sh = ss.getSheetByName(PLATFORM_POC_TAB);
  ensurePocColumns_(sh);
  var ACI  = pocAuthColIndex_(sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]);
  var salt = newSalt_();
  sh.getRange(rowIndex, ACI.pwSalt + 1).setValue(salt);
  sh.getRange(rowIndex, ACI.pwHash + 1).setValue(hashPassword_(plaintext, salt));
}

/**
 * Admin-only password provisioning, callable from the app.
 * Run setUserPasswordFromEditor() below instead when bootstrapping the very
 * first admin, since no one can sign in yet at that point.
 */
function adminSetUserPassword(token, targetEmail, newPassword) {
  try {
    var me = requireSession_(token);
    if (!me.perms.manageUsers) return { ok:false, error:'You do not have permission to manage users.' };
    var err = passwordPolicyError_(newPassword);
    if (err) return { ok:false, error:err };
    var u = findDirectoryUser_(targetEmail);
    if (!u) return { ok:false, error:'No directory row for ' + targetEmail };
    writeCredential_(u.rowIndex, newPassword);
    return { ok:true };
  } catch (e) { return { ok:false, error:String(e && e.message || e) }; }
}

/* ── One-time provisioning ────────────────────────────────────────────────────
   Run these from the Apps Script editor (Run ▸ function). They are deliberately
   not exposed to the web app: the first one creates the tab that the login
   system reads, and the second sets the first password, at which point no one
   can be signed in yet to authorise it.
   ──────────────────────────────────────────────────────────────────────────── */

/**
 * Creates the POC Directory tab with the full canonical schema if it is
 * missing, or brings an existing tab up to date by appending only the columns
 * it lacks. Never reorders, rewrites or deletes anything already there, and
 * never touches Sellers / Buyers / DocStatus / NBFC Requirements.
 */
function setupPOCDirectory() {
  var ss = SpreadsheetApp.openById(PLATFORM_SHEET_ID);
  var sh = ss.getSheetByName(PLATFORM_POC_TAB);
  var created = false;

  if (!sh) {
    sh = ss.insertSheet(PLATFORM_POC_TAB);
    sh.getRange(1, 1, 1, POC_HEADERS.length).setValues([POC_HEADERS]);
    created = true;
  }
  ensurePocColumns_(sh);

  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold').setBackground('#EEF2F7').setVerticalAlignment('middle');
  sh.setRowHeight(1, 34);

  /* Credential columns are operational plumbing, not something to read or edit
     by hand — collapse them out of the way and mark them clearly. */
  var ACI = pocAuthColIndex_(headers);
  ['pwHash','pwSalt'].forEach(function (k) {
    if (ACI[k] == null) return;
    var c = ACI[k] + 1;
    sh.getRange(1, c).setNote('Managed by the app. Never type a password here — '
      + 'use adminSetUserPassword() or Change Password in the app.');
    sh.setColumnWidth(c, 60);
  });

  Logger.log((created ? 'Created' : 'Updated') + ' "' + PLATFORM_POC_TAB + '" — '
    + headers.length + ' columns.');
  Logger.log('Columns: ' + headers.join(' | '));
  Logger.log('Next: add one row per user (POC Name, Official Email, Role, Team, '
    + 'States, Reports To, Rank, Esc L1, Esc L2, Active=Yes), then run '
    + 'setUserPasswordFromEditor() to set the first password.');
  return { ok:true, created:created, columns:headers };
}

/**
 * Bootstrap / reset one user's password from the editor.
 * Put the address and password in the two constants, run once, then CLEAR THEM
 * and save — an editor-visible plaintext password is exactly what the hashing
 * above exists to avoid, and this file is committed to git.
 */
function setUserPasswordFromEditor() {
  var EMAIL    = '';   // e.g. 'ajay.vunyale@recykal.com'
  var PASSWORD = '';   // set, run, then clear before saving

  if (!EMAIL || !PASSWORD) {
    throw new Error('Set EMAIL and PASSWORD inside setUserPasswordFromEditor(), run it, then clear them again.');
  }
  var err = passwordPolicyError_(PASSWORD);
  if (err) throw new Error(err);
  var u = findDirectoryUser_(EMAIL);
  if (!u) throw new Error('No row in "' + PLATFORM_POC_TAB + '" with Official Email = ' + EMAIL);
  writeCredential_(u.rowIndex, PASSWORD);
  Logger.log('Password set for ' + EMAIL + ' (row ' + u.rowIndex + '). Now clear the constants above.');
}

/** Lists who can sign in and what they will see. Read-only diagnostic. */
function auditAccess() {
  var ss = SpreadsheetApp.openById(PLATFORM_SHEET_ID);
  var sh = ss.getSheetByName(PLATFORM_POC_TAB);
  if (!sh) { Logger.log('No "' + PLATFORM_POC_TAB + '" tab — run setupPOCDirectory() first.'); return; }
  var vals = sh.getDataRange().getValues();
  if (vals.length < 2) { Logger.log('Directory is empty.'); return; }
  var CI = pocColIndex_(vals[0]), ACI = pocAuthColIndex_(vals[0]);
  Logger.log('email | role → tier | active | password set | states');
  for (var r = 1; r < vals.length; r++) {
    function f(i) { return i == null ? '' : String(vals[r][i] == null ? '' : vals[r][i]).trim(); }
    var em = f(CI.email); if (!em) continue;
    var rk = roleKey_(f(CI.role));
    Logger.log([em, f(CI.role) || '(blank)' + ' → ' + rk, isActiveFlag_(f(CI.active)) ? 'ACTIVE' : 'denied',
      f(ACI.pwHash) ? 'yes' : 'NO', ROLE_PERMISSIONS[rk].viewAllStates ? 'ALL' : (f(CI.states) || '(none)')
    ].join(' | '));
  }
}

/* ── Scoping ──────────────────────────────────────────────────────────────── */

/**
 * Removes rows the caller is not entitled to see. Applied before the response
 * is assembled, so restricted records are never sent to the browser at all.
 * Admin sees everything; anyone else is limited to their assigned States, and
 * a user with no states assigned sees nothing rather than everything.
 */
function applyScope_(me, rows) {
  if (me.perms.viewAllStates) return rows;
  var allowed = me.states || [];
  if (!allowed.length) return [];                       // fail closed
  var out = rows.filter(function (r) {
    return allowed.indexOf(String(r.state || '').trim().toLowerCase()) >= 0;
  });
  /* Team narrowing is applied only where the row actually carries a team
     marker. The Sellers/Buyers tabs currently leave the Type column blank on
     every row, so this is a no-op until that column is populated — filtering on
     an empty column would silently hide the entire book. */
  if (!me.perms.viewAllTeams && me.team) {
    var team = String(me.team).trim().toLowerCase();
    var typed = out.filter(function (r) { return String(r.businessType || '').trim(); });
    if (typed.length) {
      out = out.filter(function (r) {
        var bt = String(r.businessType || '').trim().toLowerCase();
        return !bt || bt === team || team.indexOf(bt) >= 0 || bt.indexOf(team) >= 0;
      });
    }
  }
  return out;
}

/** Read POC Directory + Followups → { pocs:[…], followups:{ pocId:[…] } }. */
function getPOCData(token) {
  try {
    var me = requireSession_(token);
    var ss = SpreadsheetApp.openById(PLATFORM_SHEET_ID);
    var pocs = [], followups = {};

    var sh = ss.getSheetByName(PLATFORM_POC_TAB);
    if (sh) {
      var v = sh.getDataRange().getValues();
      if (v.length >= 2) {
        var CI = pocColIndex_(v[0]);
        function gv(row, key) {
          var i = CI[key]; if (i == null) return '';
          var val = row[i];
          if (val instanceof Date) return val.toISOString();
          return String(val == null ? '' : val).trim();
        }
        for (var r = 1; r < v.length; r++) {
          var id = gv(v[r], 'id');
          if (!id) continue;
          pocs.push({
            id:            id,
            nbfc:          gv(v[r], 'nbfc'),
            entityType:    gv(v[r], 'entityType'),
            linkedSeller:  gv(v[r], 'linkedSeller'),
            linkedBuyer:   gv(v[r], 'linkedBuyer'),
            org:           gv(v[r], 'org'),
            name:          gv(v[r], 'name'),
            designation:   gv(v[r], 'designation'),
            department:    gv(v[r], 'department'),
            mobile:        gv(v[r], 'mobile'),
            altMobile:     gv(v[r], 'altMobile'),
            email:         gv(v[r], 'email'),
            altEmail:      gv(v[r], 'altEmail'),
            whatsapp:      gv(v[r], 'whatsapp'),
            channel:       gv(v[r], 'channel'),
            primary:       /^(yes|true|1|primary)$/i.test(gv(v[r], 'primary')),
            role:          gv(v[r], 'role'),
            relOwner:      gv(v[r], 'relOwner'),
            status:        gv(v[r], 'status') || 'Active',
            lastContact:   gv(v[r], 'lastContact'),
            nextFollowup:  gv(v[r], 'nextFollowup'),
            followupStatus:gv(v[r], 'followupStatus'),
            remarks:       gv(v[r], 'remarks'),
            createdBy:     gv(v[r], 'createdBy'),
            createdOn:     gv(v[r], 'createdOn'),
            modifiedBy:    gv(v[r], 'modifiedBy'),
            modifiedOn:    gv(v[r], 'modifiedOn'),
            /* Access attributes. Directory data, not credentials — the hash and
               salt columns are absent from POC_FIELDS and so cannot be read by
               gv() at all. Escalation addresses are served live from here so the
               app never carries a hardcoded escalation contact. */
            team:          gv(v[r], 'team'),
            states:        gv(v[r], 'states'),
            reportsTo:     gv(v[r], 'reportsTo'),
            rank:          gv(v[r], 'rank'),
            escL1:         gv(v[r], 'escL1'),
            escL2:         gv(v[r], 'escL2'),
            active:        /^(yes|y|true|1|active)$/i.test(gv(v[r], 'active'))
          });
        }
      }
    }

    var fsh = ss.getSheetByName(PLATFORM_POCFU_TAB);
    if (fsh) {
      var fv = fsh.getDataRange().getValues();
      if (fv.length >= 2) {
        var FH = {}; fv[0].forEach(function (h, i) { FH[String(h).trim().toLowerCase()] = i; });
        function fc(row, name) { var i = FH[name.toLowerCase()]; return i == null ? '' : String(row[i] == null ? '' : row[i]).trim(); }
        for (var d = 1; d < fv.length; d++) {
          var pid = fc(fv[d], 'poc id');
          if (!pid) continue;
          (followups[pid] = followups[pid] || []).push({
            id: fc(fv[d], 'followup id'), pocId: pid, nbfc: fc(fv[d], 'nbfc'), date: fc(fv[d], 'date'),
            mode: fc(fv[d], 'mode'), summary: fc(fv[d], 'summary'), nextAction: fc(fv[d], 'next action'),
            nextDate: fc(fv[d], 'next date'), by: fc(fv[d], 'by')
          });
        }
      }
    }
    /* scrubClientPayload_ is the backstop: even if a credential-shaped column is
       ever added to POC_FIELDS by mistake, it cannot reach the browser. */
    return { ok: true, pocs: scrubClientPayload_(pocs), followups: followups,
             generatedAt: new Date().toISOString() };
  } catch (e) {
    var msg = String(e && e.message || e);
    if (msg === 'AUTH_REQUIRED') {
      return { ok:false, authRequired:true, error:'Your session has expired. Please sign in again.',
               pocs:[], followups:{} };
    }
    return { ok: false, error: msg, pocs: [], followups: {} };
  }
}

/** Upsert one contact (keyed by the permanent POC ID — never by name/email/phone).
    The canonical schema (POC_FIELDS) drives the write, missing columns are added
    on the fly, and an audit trail is maintained: Created By/On are stamped once on
    insert and preserved on every later update; Last Modified By/On are refreshed on
    each write. Entity mapping is validated (a seller NBFC only accepts a seller POC,
    a buyer NBFC only a buyer POC). A script lock serialises concurrent writes. */
function savePOC(poc, token) {
  var me;
  try {
    me = requireSession_(token);
    if (!me.perms.managePOC) return { ok:false, error:'Your role cannot edit POC records.' };
  } catch (e) { return { ok:false, authRequired:true, error:'Your session has expired. Please sign in again.' }; }
  var lock = LockService.getScriptLock();
  try { lock.waitLock(20000); } catch (e) { return { ok: false, error: 'POC Directory is busy, please retry.' }; }
  try {
    poc = poc || {};
    var id = String(poc.id || '').trim();
    if (!id) return { ok: false, error: 'POC ID is required' };

    // Entity-mapping guard — reject a POC whose entity type contradicts its NBFC group.
    var grp = nbfcEntityGroup_(poc.nbfc);
    var et  = String(poc.entityType || '').trim().toLowerCase();
    if (grp && et && grp !== et) {
      return { ok: false, error: 'Invalid mapping: "' + poc.nbfc + '" is a ' + grp +
        ' NBFC and cannot hold a ' + et + ' POC.' };
    }

    var ss = SpreadsheetApp.openById(PLATFORM_SHEET_ID);
    var sh = ss.getSheetByName(PLATFORM_POC_TAB);
    if (!sh) {
      sh = ss.insertSheet(PLATFORM_POC_TAB);
      sh.getRange(1, 1, 1, POC_HEADERS.length).setValues([POC_HEADERS]);
      sh.setFrozenRows(1);
    }
    var CI   = ensurePocColumns_(sh);        // guarantees every canonical column exists
    var vals = sh.getDataRange().getValues();
    var now  = new Date();
    var actor = String(poc.modifiedBy || poc.createdBy || poc.by || '').trim();

    var rowNum = -1;
    if (CI.id != null) {
      for (var r = 1; r < vals.length; r++) { if (String(vals[r][CI.id]).trim() === id) { rowNum = r + 1; break; } }
    }
    function setCell(key, value) { if (CI[key] != null) sh.getRange(rowNum, CI[key] + 1).setValue(value); }

    if (rowNum > 0) {
      // Update — write every mapped field EXCEPT the create-audit pair (preserved).
      POC_FIELDS.forEach(function (f) {
        if (f.key === 'createdBy' || f.key === 'createdOn') return;   // keep original creator
        if (f.key === 'modifiedBy' || f.key === 'modifiedOn') return; // stamped just below
        if (CI[f.key] != null) setCell(f.key, pocFieldValue_(poc, f.key));
      });
      setCell('modifiedBy', actor);
      setCell('modifiedOn', now);
    } else {
      // Insert — full canonical row, stamping both create and modify audit fields.
      var rec = {}; for (var k in poc) if (poc.hasOwnProperty(k)) rec[k] = poc[k];
      rec.createdBy  = actor; rec.createdOn  = now;
      rec.modifiedBy = actor; rec.modifiedOn = now;
      var width  = Math.max(sh.getLastColumn(), POC_HEADERS.length);
      var newRow = new Array(width).fill('');
      POC_FIELDS.forEach(function (f) { if (CI[f.key] != null) newRow[CI[f.key]] = pocFieldValue_(rec, f.key); });
      sh.appendRow(newRow);
    }

    // Enforce one primary POC per NBFC.
    if (poc.primary && CI.primary != null && CI.nbfc != null && CI.id != null) {
      var v2 = sh.getDataRange().getValues();
      for (var k2 = 1; k2 < v2.length; k2++) {
        if (String(v2[k2][CI.nbfc]).trim() === String(poc.nbfc).trim() && String(v2[k2][CI.id]).trim() !== id &&
            /^(yes|true|1|primary)$/i.test(String(v2[k2][CI.primary]))) {
          sh.getRange(k2 + 1, CI.primary + 1).setValue('No');
        }
      }
    }
    return { ok: true, id: id };
  } catch (e) {
    return { ok: false, error: String(e && e.message || e) };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/** Append one interaction to the Followups tab. */
function savePOCFollowup(fu, token) {
  try {
    var me = requireSession_(token);
    if (!me.perms.managePOC) return { ok:false, error:'Your role cannot log follow-ups.' };
    fu = fu || {};
    if (!fu.pocId) return { ok: false, error: 'pocId is required' };
    var ss = SpreadsheetApp.openById(PLATFORM_SHEET_ID);
    var sh = ss.getSheetByName(PLATFORM_POCFU_TAB);
    if (!sh) { sh = ss.insertSheet(PLATFORM_POCFU_TAB); sh.getRange(1, 1, 1, POCFU_HEADERS.length).setValues([POCFU_HEADERS]); sh.setFrozenRows(1); }
    sh.appendRow([fu.id || ('FU-' + fu.pocId + '-' + Date.now()), fu.pocId, fu.nbfc || '', fu.date || '',
                  fu.mode || '', fu.summary || '', fu.nextAction || '', fu.nextDate || '', fu.by || '', new Date()]);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: String(e && e.message || e) };
  }
}

/* ═══════════════════════════════════════════════════════════════════════════
   NBFC REQUIREMENT MATRIX  ·  backend-configurable document requirements
   ───────────────────────────────────────────────────────────────────────────
   One row per NBFC in the "NBFC Requirements" tab holds its mandatory document
   ids (comma-separated) and an optional vintage threshold. Editing that row —
   from the sheet or the dashboard's Settings → NBFC Requirements — changes what
   the Needs-documents panel and the eligibility engine require, with no code
   change. Never touched by the Metabase sync.
   ═══════════════════════════════════════════════════════════════════════════ */
var PLATFORM_REQ_TAB = 'NBFC Requirements';

/** Read the matrix → { config: { <nbfcId>: { req:[docId…], vintageYrs } } }. */
function getNbfcConfig() {
  try {
    var ss = SpreadsheetApp.openById(PLATFORM_SHEET_ID);
    var sh = ss.getSheetByName(PLATFORM_REQ_TAB);
    var config = {};
    if (sh) {
      var v = sh.getDataRange().getValues();
      if (v.length >= 2) {
        var H = {}; v[0].forEach(function (h, i) { H[String(h).trim().toLowerCase()] = i; });
        var gi = H['nbfc id'], di = H['required doc ids'], vi = H['vintage years'];
        for (var r = 1; r < v.length; r++) {
          var id = gi != null ? String(v[r][gi]).trim() : '';
          if (!id) continue;
          var raw = di != null ? String(v[r][di] == null ? '' : v[r][di]) : '';
          var req = raw.split(',').map(function (s) { return s.trim(); }).filter(String);
          var vy = vi != null ? String(v[r][vi] == null ? '' : v[r][vi]).trim() : '';
          config[id] = { req: req, vintageYrs: vy === '' ? '' : Number(vy) };
        }
      }
    }
    return { ok: true, config: config };
  } catch (e) {
    return { ok: false, error: String(e && e.message || e), config: {} };
  }
}

/** Upsert one NBFC's requirement row. payload = { id, req:[…], vintageYrs }. */
function saveNbfcConfig(payload, token) {
  try {
    var me = requireSession_(token);
    if (!me.perms.configure) return { ok:false, error:'Only an Admin can change NBFC requirements.' };
    payload = payload || {};
    var id = String(payload.id || '').trim();
    if (!id) return { ok: false, error: 'NBFC id is required' };
    var req = (payload.req || []).join(',');
    var vy  = (payload.vintageYrs == null ? '' : payload.vintageYrs);
    var ss = SpreadsheetApp.openById(PLATFORM_SHEET_ID);
    var sh = ss.getSheetByName(PLATFORM_REQ_TAB);
    if (!sh) { sh = ss.insertSheet(PLATFORM_REQ_TAB); sh.getRange(1, 1, 1, 4).setValues([['NBFC ID', 'Required Doc IDs', 'Vintage Years', 'Updated At']]); sh.setFrozenRows(1); }
    var vals = sh.getDataRange().getValues(), rowNum = -1;
    for (var r = 1; r < vals.length; r++) { if (String(vals[r][0]).trim() === id) { rowNum = r + 1; break; } }
    var row = [id, req, vy, new Date()];
    if (rowNum > 0) sh.getRange(rowNum, 1, 1, 4).setValues([row]);
    else            sh.appendRow(row);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: String(e && e.message || e) };
  }
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
