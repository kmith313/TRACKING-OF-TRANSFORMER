/*****************************************************************
 *  TRANSFORMER TRACKING APP  -  Google Apps Script (backend)
 *  Each organisation gives its Google Sheet link to the app
 *  (Share -> "Anyone with the link" -> Editor). On setup, these tabs are created in the Sheet:
 *  DATA, Users (Admin), Drivers (Driver Master), Config, ARCHIVE.
 *****************************************************************/
var SS_ID = '';                                         // set on every request from the token / login
var TZ = 'Asia/Kolkata';
var DATA = 'DATA', USERS = 'Users', DRV = 'Drivers', ARC = 'ARCHIVE', CONF = 'Config', KIT = 'Kitty';
var ARCH_YEARS = 2;                                     // (only used when ARCH_FY = false, and for "replaced but never returned" records) years after which a record may be archived
var MAX_WS = 12;                                        // maximum workshops per app

// DATA sheet columns - only up to REPLACEMENT DATE + TIME (pasted in this order)
var COLS = ['Workshop','DIVISION','SUBSTATION','FEEDER','PLACE OF DAMAGE','CAPACITY','Category of Transformer','COMPLAIN NUMBER','PR NO','PR DATE','JE Name','JE Mobile','ISSUED TO FIRM','ISSUE DATE','DRIVER NAME','DRIVER MOBILE','REPLACEMENT DATE','TIME','RETURN DATE','REMARK'];
var NC = COLS.length;                                   // 20 (last = REMARK: the JE writes why a transformer is not replaced yet)
var DATE_IDX = [9,13,16,18];                               // columns that hold a DD.MM.YYYY date
var I_WS=0, I_DIV=1, I_SUB=2, I_CAP=5, I_PR=8, I_PRD=9, I_ISS=13, I_REP=16, I_TIME=17, I_RET=18, EDIT_FROM=10;
var STD_CAPS = [10,16,25,63,100,160,250,400,630];
// A wrong / duplicate PR is marked by choosing the default firm "DUPLICATE/WRONG PR" (present in every Driver Master).
// Such a record has no issue / replacement / return date and is not counted as damage: the summary shows it in its own column.
var DUP_FIRM = 'DUPLICATE/WRONG PR', DUP_DRV = 'JE(D)', I_FIRM = 12, I_REMARK = 19;
function isDup_(firm) { return /^DUPLICATE\s*\/?\s*WRONG\s*PR\.?$/i.test(String(firm || '').replace(/\s+/g, ' ').trim()); }
function dupFix_(rec) {                                  // wrong / duplicate PR: no issue, replacement or return
  if (!isDup_(rec[I_FIRM])) return rec;
  rec[I_FIRM] = DUP_FIRM; rec[I_ISS] = ''; rec[I_REP] = ''; rec[I_TIME] = ''; rec[I_RET] = ''; return rec;
}

/* ---------------- web app entry (JSON API - frontend GitHub Pages par hai) ---------------- */
// The developer deploys this script ONCE:  Deploy -> New deployment -> Web app
//   Execute as: Me   |   Who has access: Anyone
// Then put the resulting /exec URL in the frontend's config.js. Every organisation's Google Sheet
// (Share -> Anyone with the link -> Editor) is read / written through this script.
var G_ = this;
var API = {   // only these functions can be called from outside
  checkSheet: 1, setupApp: 1, login: 1, getMe: 1, logout: 1, changePassword: 1,
  adminUsers: 1, adminSetScope: 1, adminRenameWorkshop: 1, adminDeleteWorkshop: 1, adminStorage: 1, adminMigrateCheck: 1, adminMigrate: 1, adminReset: 1, adminSetActive: 1, adminAddUser: 1, adminAddWorkshop: 1, adminSetDivision: 1, adminBackupInfo: 1, adminRestoreOtp: 1, adminRestore: 1, restoreNewSheet: 1, creatorBackups: 1, creatorBackupOn: 1, creatorBackupNow: 1, getDiscomSummary: 1, dhDivs: 1, dhPick: 1, creatorSetTabs: 1, creatorLogin: 1, creatorHeads: 1, creatorAddHead: 1, creatorResetHead: 1, creatorSetActive: 1, adminInfo: 1, adminArchive: 1, adminConvert: 1, adminHealth: 1,
  getData: 1, saveGrid: 1, saveRows: 1, deleteRecord: 1, saveRecord: 1,
  getSummary: 1, getKitty: 1, getKittyPrev: 1, kitSavePrev: 1, kitRequestOpen: 1, kitCancelReq: 1, kitPending: 1, kitDecide: 1, saveKitty: 1, getCompare: 1, getDiscomCompare: 1, getList: 1, getCols: 1, getDataRange: 1,
  getDrivers: 1, saveDrivers: 1, deleteDrivers: 1,
  discomList: 1, complainSearch: 1, complainIndex: 1, complainVersion: 1, complainDetail: 1, adminForgot: 1            // public (no login): complaint status page
};
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doGet() { return json_({ ok: true, app: 'Transformer Tracking API', time: new Date().toISOString() }); }
var CUR_FN = '';                                          // name of the function being called (a DISCOM head may only reach the tabs the Creator allowed)
function doPost(e) {
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!req.fn || !API[req.fn] || typeof G_[req.fn] !== 'function') return json_({ ok: false, e: 'Unknown function' });
    CUR_FN = req.fn;
    var out = G_[req.fn].apply(null, req.args || []);
    return json_({ ok: true, v: out === undefined ? null : out });
  } catch (err) {
    return json_({ ok: false, e: String((err && err.message) || err) });
  }
}

/* ---------------- small helpers ---------------- */
var SS_MEM = null;
function useSheet_(id) {                                // the sheet of this request
  if (SS_ID !== id) { SS_ID = id; SS_MEM = null; IDX_MEM = null; ARCH_IX = null; CFG_MEM = null; }
}
function ss_() {                                        // open once, not repeatedly (each openById takes ~0.3 sec)
  if (!SS_ID) throw new Error('No Google Sheet is connected.');
  if (SS_MEM) return SS_MEM;
  SS_MEM = SpreadsheetApp.openById(SS_ID);
  try { ensureCols_(SS_MEM); } catch (e) {}
  return SS_MEM;
}
function ensureCols_(ss) {                              // older sheets: add the newer columns (RETURN DATE, REMARK) once
  var sh = ss.getSheetByName(DATA); if (!sh) return;
  var cc = CacheService.getScriptCache(), fk = 'hdr' + NC + '_' + SS_ID;
  if (cc.get(fk)) return;
  if (sh.getMaxColumns() < NC) sh.insertColumnsAfter(sh.getMaxColumns(), NC - sh.getMaxColumns());
  var hd = sh.getRange(1, 1, 1, NC).getValues()[0], fixed = false;
  for (var i = 0; i < NC; i++) if (String(hd[i]).trim().toUpperCase() !== COLS[i].toUpperCase() && !String(hd[i]).trim()) {
    sh.getRange(1, i + 1).setValue(COLS[i]).setFontWeight('bold').setBackground('#e3e9fb');
    sh.getRange(1, i + 1, sh.getMaxRows(), 1).setNumberFormat('@'); sh.getRange(1, i + 1).setValue(COLS[i]);
    sh.setColumnWidth(i + 1, i === NC - 1 ? 260 : 120); fixed = true;
  }
  if (fixed) bump_();
  cc.put(fk, '1', 21600);
}
// link or plain ID -> Sheet ID
function parseId_(s) {
  s = String(s || '').trim();
  var m = /\/d\/([a-zA-Z0-9_-]{20,})/.exec(s) || /^([a-zA-Z0-9_-]{25,})$/.exec(s);
  return m ? m[1] : '';
}
var NOSHEET = 'The Google Sheet could not be opened. Please check: (1) the link is complete and correct, (2) in the Sheet: Share -> General access -> "Anyone with the link" -> role "Editor".';
function openSafe_(id) { try { return SpreadsheetApp.openById(id); } catch (e) { throw new Error(NOSHEET); } }
// getUi() does not work when run from the editor - then it logs instead
function alert_(m) { try { SpreadsheetApp.getUi().alert(m); } catch (e) { Logger.log(m); } }
var MONTHS = ['JAN','FEB','MAR','APR','MAY','JUN','JUL','AUG','SEP','OCT','NOV','DEC'];
function rnd_(n) {
  var ch = 'abcdefghjkmnpqrstuvwxyz23456789', s = '', b = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  for (var i = 0; i < n; i++) s += ch.charAt(parseInt(b.substr(i * 2, 2), 16) % ch.length);
  return s;
}
function bool_(v) { return v === true || String(v).toUpperCase() === 'TRUE'; }
function p2_(n) { return ('0' + n).slice(-2); }
/* ---------------- config (workshop list) - Config tab me rehti hai ---------------- */
var CFG_MEM = null;
function cfgRead_() {
  var sh = ss_().getSheetByName(CONF), o = { org: '', discom: '', division: '', workshops: [], old: [], movedTo: '', conv: false, ym: '' };
  if (!sh || sh.getLastRow() < 1) return o;
  sh.getRange(1, 1, sh.getLastRow(), 2).getValues().forEach(function (r) {
    var k = String(r[0]).trim(), v = String(r[1] == null ? '' : r[1]).trim();
    if (k === 'Org') o.org = v;
    if (k === 'Discom') o.discom = v;
    if (k === 'Division') o.division = v;
    if (k === 'MovedTo') o.movedTo = v;
    if (k === 'Converted') o.conv = (v === '1');
    if (k === 'ArchYM') o.ym = v;
    if (k === 'OldSheets') { try { o.old = JSON.parse(v) || []; } catch (e) { o.old = []; } }
    if (k === 'Workshops') o.workshops = v.split(',').map(function (x) { return x.replace(/\s+/g, ' ').trim().toUpperCase(); }).filter(Boolean);
  });
  return o;
}
function cfgSet_(key, val) {                                // write one setting in the Config tab
  var sh = ss_().getSheetByName(CONF), v = sh.getLastRow() ? sh.getRange(1, 1, sh.getLastRow(), 1).getValues() : [], row = 0;
  for (var i = 0; i < v.length; i++) if (String(v[i][0]).trim() === key) { row = i + 1; break; }
  if (!row) { row = sh.getLastRow() + 1; sh.getRange(row, 1).setValue(key); }
  sh.getRange(row, 2).setNumberFormat('@').setValue(val);
}
var CFG_T = 0;
function cfg_() {
  if (CFG_MEM && Date.now() - CFG_T < 60000) return CFG_MEM;          // a warm server instance forgets its copy after 1 minute (another instance may have changed a setting)
  CFG_T = Date.now();
  var c = CacheService.getScriptCache(), k = 'cfg_' + SS_ID, s = c.get(k);
  if (s) { try { return (CFG_MEM = JSON.parse(s)); } catch (e) {} }
  var o = cfgRead_();
  if (o.workshops.length) { try { c.put(k, JSON.stringify(o), 1800); } catch (e) {} }
  return (CFG_MEM = o);
}
function cfgSetQ_(key, val) { cfgSet_(key, val); CFG_MEM = null; try { CacheService.getScriptCache().remove('cfg_' + SS_ID); } catch (e) {} }   // a setting that does not touch the data (no index rebuild)
function cfgBump_() { CFG_MEM = null; try { CacheService.getScriptCache().remove('cfg_' + SS_ID); } catch (e) {} bump_(); }
function ws_() { return cfg_().workshops; }                               // ["AZAMGARH","MAU",...]
function wsAt_(i) { return i >= 0 ? (ws_()[i] || '') : ''; }
// any written workshop name -> the canonical name from the list ('' if not recognised)
function nkw_(s) { var t = nk_(s); return /^EW[CDS]./.test(t) && t.length > 5 ? t.slice(3) : t; }   // "EWC AZAMGARH" / "EWD AZAMGARH" -> "AZAMGARH"
function wsKey_(s) {
  var t = nk_(s); if (!t) return '';
  var W = ws_(), i, t2 = nkw_(s);
  for (i = 0; i < W.length; i++) if (nk_(W[i]) === t) return W[i];
  for (i = 0; i < W.length; i++) if (nkw_(W[i]) === t2) return W[i];
  var hit = [];                                                           // one contained in the other (e.g. "EWC AZAMGARH") - only when exactly one matches
  for (i = 0; i < W.length; i++) { var w = nkw_(W[i]); if (w.length >= 3 && (t2.indexOf(w) >= 0 || (t2.length >= 3 && w.indexOf(t2) >= 0))) hit.push(W[i]); }
  return hit.length === 1 ? hit[0] : '';
}
function wsName_(v) {
  var n = String(v == null ? '' : v).replace(/\s+/g, ' ').trim().toUpperCase();
  if (!/^[A-Z0-9][A-Z0-9 .&()\-]{1,29}$/.test(n)) throw new Error('Workshop name "' + n + '" is not valid: 2-30 characters, only A-Z, 0-9, space and . & ( ) -');
  return n;
}
function divKey_(s) { return String(s || '').replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim().toUpperCase(); }
// "25 KVA 11/0.4 KV TRANSFORMER REPAIRED" -> 25
function capOf_(s) {
  var t = String(s || '').replace(/\u00a0/g, ' ').trim();
  var U = t.toUpperCase(), m = U.match(/(?:^|[^\d.\/])(\d+)\s*K\.?\s*V\.?\s*A/) || U.match(/(?:^|[^\d.\/])(\d+)\s*KV/);   // "25 KVA 11/0.4 KV ..." -> 25 (KVA first, voltage is ignored)
  if (m) return +m[1];
  if (/^\d+(\.0+)?$/.test(t)) return parseInt(t, 10);   // also when only "25" is written
  return 0;
}
// any date -> "DD.MM.YYYY"
function toDMY_(v) {
  if (v === null || v === undefined || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return isNaN(v) ? '' : Utilities.formatDate(v, TZ, 'dd.MM.yyyy');
  }
  if (typeof v === 'number') {
    if (v > 20000 && v < 80000) return Utilities.formatDate(new Date(Math.round((v - 25569) * 86400000)), 'UTC', 'dd.MM.yyyy');
    return String(v);
  }
  var s = String(v).trim(); if (!s) return '';
  var m = s.match(/^(\d{1,2})([.\/\-])(\d{1,2})[.\/\-](\d{2,4})/);
  if (m) {                                                  // 18.10.2026 / 18-10-2026 = DD.MM ; 10/18/2026 (ERP) = MM/DD  (a slash date is read as MM/DD/YYYY, unless the first number is above 12)
    var y = +m[4], a1 = +m[1], b1 = +m[3], dd = a1, mm = b1; if (y < 100) y += 2000;
    if (m[2] === '/' && a1 <= 12) { mm = a1; dd = b1; }
    return p2_(dd) + '.' + p2_(mm) + '.' + y;
  }
  m = s.match(/^(\d{4})[.\/\-](\d{1,2})[.\/\-](\d{1,2})/);
  if (m) return p2_(m[3]) + '.' + p2_(m[2]) + '.' + m[1];
  m = s.match(/^(\d{1,2})[\-\s\/.]([A-Za-z]{3,9})[\-\s\/.,]*(\d{2,4})/);      // 14-Oct-26 / 14 October 2026
  if (m) {
    var mi = MONTHS.indexOf(m[2].slice(0, 3).toUpperCase());
    if (mi >= 0) { var y2 = +m[3]; if (y2 < 100) y2 += 2000; return p2_(m[1]) + '.' + p2_(mi + 1) + '.' + y2; }
  }
  return s;
}
// "DD.MM.YYYY" -> 20261014 (for comparison), invalid/blank = 0
function dk_(s) {
  var m = /^(\d{2})\.(\d{2})\.(\d{4})$/.exec(String(s || ''));
  if (!m) return 0;
  var d = +m[1], mo = +m[2], y = +m[3];
  if (y < 1990 || y > 2100 || mo < 1 || mo > 12 || d < 1 || d > 31) return 0;
  var t = new Date(y, mo - 1, d);
  if (t.getMonth() !== mo - 1 || t.getDate() !== d) return 0;      // 31.02.2026 jaisi date
  return y * 10000 + mo * 100 + d;
}
function norm_(i, v) {
  if (v === null || v === undefined) return '';
  if (i === I_TIME && Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, TZ, 'HH:mm');
  if (DATE_IDX.indexOf(i) >= 0) return toDMY_(v);
  if (i === I_TIME && typeof v === 'number' && v >= 0 && v < 1) {
    var mins = Math.round(v * 1440); return p2_(Math.floor(mins / 60) % 24) + ':' + p2_(mins % 60);
  }
  var s = String(v).replace(/\u00a0/g, ' ').trim();
  if (i === I_PR && /^\d+\.0+$/.test(s)) s = s.replace(/\.0+$/, '');
  return s;
}
var MOB_IDX = [11, 15];                                  // JE Mobile, DRIVER MOBILE
// returns a 10-digit number; '' if empty; null if invalid (E+ notation, too short/long)
function mobClean_(v) {
  var s = String(v == null ? '' : v).replace(/\u00a0/g, ' ').trim();
  if (!s) return '';
  var m = /^(\d)(?:\.(\d+))?\s*e\s*\+?(\d+)$/i.exec(s);                 // 8.887812345E+09 jaisa
  if (m) { var dg = m[1] + (m[2] || ''), ex = +m[3] + 1; if (dg.length !== ex) return null; s = dg; }
  var d = s.replace(/\D/g, '');
  if (d.length === 12 && d.indexOf('91') === 0) d = d.slice(2); else if (d.length === 11 && d.charAt(0) === '0') d = d.slice(1);
  return d.length === 10 ? d : null;
}
function lock_() { var l = LockService.getScriptLock(); try { l.waitLock(25000); } catch (e) { throw new Error('The server is busy (lock timeout), please try again.'); } return l; }

/* ---------------- users / login ---------------- */
function hash_(pw, salt) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + pw)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}
function mk_(pw) { var s = Utilities.getUuid().slice(0, 8); return 'h1$' + s + '$' + hash_(pw, s); }
function chk_(stored, pw) {
  if (stored.indexOf('h1$') === 0) { var p = stored.split('$'); return hash_(pw, p[1]) === p[2]; }
  return stored === pw;
}
function usersBump_() { try { CacheService.getScriptCache().remove('us_' + SS_ID); } catch (e) {} }
function users_() {                                       // cached 5 min (cleared on every change) - no need to open the Sheet for every request
  var cc = CacheService.getScriptCache(), ck = 'us_' + SS_ID, hit = cc.get(ck);
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  var out = usersRead_();
  try { cc.put(ck, JSON.stringify(out), 300); } catch (e) {}
  return out;
}
function usersRead_() {
  var sh = ss_().getSheetByName(USERS), n = sh.getLastRow(), out = [];
  if (n < 2) return out;
  sh.getRange(2, 1, n - 1, 9).getValues().forEach(function (r, i) {
    var name = String(r[0]).trim(); if (!name) return;
    var role = String(r[2] || '').trim();
    var scope = String(r[6] || '').trim().toUpperCase();
    if (!scope) scope = role.toUpperCase().indexOf('AE/JE') === 0 ? (wsKey_(name) || 'NONE') : 'ALL';
    out.push({
      row: i + 2, name: name, key: name.toLowerCase(), pass: String(r[1]), role: role,
      display: String(r[3] || name), canEdit: bool_(r[4]), active: bool_(r[5]),
      scope: scope.split(',').map(function (x) { return x.trim(); }).filter(Boolean), admin: bool_(r[7]), must: bool_(r[8])
    });
  });
  return out;
}
var SHEET_LIMIT = 20000000;                              // Google Sheets: 20 million cells per file
function usage_() {
  var c = CacheService.getScriptCache(), k = 'use_' + SS_ID, h = c.get(k);
  if (h) { try { return JSON.parse(h); } catch (e) {} }
  var cells = 0; ss_().getSheets().forEach(function (sh) { cells += sh.getMaxRows() * sh.getMaxColumns(); });
  var o = { cells: cells, limit: SHEET_LIMIT, pct: Math.round(cells * 1000 / SHEET_LIMIT) / 10 };
  try { c.put(k, JSON.stringify(o), 600); } catch (e) {}
  return o;
}
function me_(u) {
  var ux = null; if (u.admin) { try { ux = usage_(); } catch (e) { ux = null; } }
  return { usage: ux, old: (cfg_().old || []).length, name: u.name, display: u.display, role: u.role, canEdit: u.canEdit, admin: u.admin, must: u.must, scope: u.scope, wsAll: ws_(), org: cfg_().org, discom: cfg_().discom, division: cfg_().division || '' };
}
function setPass_(u, pw, must) {
  var sh = ss_().getSheetByName(USERS);
  sh.getRange(u.row, 2).setValue(mk_(pw));
  sh.getRange(u.row, 9).setValue(!!must);
  usersBump_();
}
function adminStorage(token) {
  var a = auth_(token); if (!a.admin) throw new Error('Admin only.');
  try { CacheService.getScriptCache().remove('use_' + SS_ID); } catch (e) {}
  return { usage: usage_(), old: cfg_().old || [], sheet: 'https://docs.google.com/spreadsheets/d/' + SS_ID + '/edit' };
}
function auth_(token, allowMust) {
  var raw = token ? CacheService.getScriptCache().get('s_' + token) : null, se = null;
  if (raw) { try { se = JSON.parse(raw); } catch (e) { se = null; } }
  if (se && se.dh) return dhView_(token, se);                // DISCOM head looking at one division (read only)
  if (!se || !se.s || !se.k) throw new Error('SESSION');
  useSheet_(se.s);
  if (cfg_().movedTo) throw new Error('SESSION');            // this sheet was replaced by a new one: log in again
  var u = users_().filter(function (x) { return x.key === se.k; })[0];
  if (!u || !u.active) throw new Error('SESSION');
  if (u.must && !allowMust) throw new Error('MUSTCHANGE');      // the temporary password must be replaced first
  return u;
}
function login(sid, username, password, want) {
  username = String(username || '').trim().toLowerCase();
  var cache = CacheService.getScriptCache();
  var dh = !parseId_(sid) && /^dh/.test(username) && dhFind_(username); if (dh) return dhLogin_(dh, password);          // DISCOM head
  var cands = parseId_(sid) ? [parseId_(sid)] : (regUserSids_(username).length ? regUserSids_(username) : []);
  sid = cands[0] || '';
  var lk = 'fail_' + (sid || 'none') + '_' + username.slice(0, 40);
  if (+(cache.get(lk) || 0) >= 5) return { ok: false, msg: 'Too many wrong passwords. Try again after 10 minutes.' };
  if (!sid) { cache.put(lk, String(+(cache.get(lk) || 0) + 1), 600); return { ok: false, msg: 'The ID or password is wrong.' }; }
  var u = null, anySheet = false, setupMiss = false;
  for (var ci = 0; ci < cands.length && !u; ci++) {                  // a transferred user has the same ID in two organisations: the password decides which one opens
    try { useSheet_(cands[ci]); if (!cache.get('us_' + cands[ci]) && !ss_().getSheetByName(USERS)) { setupMiss = true; continue; } anySheet = true; }
    catch (e) { continue; }
    if (cfg_().movedTo) continue;
    var f = users_().filter(function (x) { return x.key === username; })[0];
    if (f && f.active && chk_(f.pass, String(password || ''))) { u = f; sid = cands[ci]; }
  }
  if (!u && !anySheet) return { ok: false, msg: setupMiss ? 'This organisation has not been set up. Ask the Admin to set it up.' : NOSHEET };
  if (!u) {
    cache.put(lk, String(+(cache.get(lk) || 0) + 1), 600);
    return { ok: false, msg: 'The ID or password is wrong.' };
  }
  cache.remove(lk);
  try { monthRoll_(); } catch (e) { Logger.log('monthRoll: ' + e); }
  if (u.pass.indexOf('h1$') !== 0) setPass_(u, String(password), u.must);
  var t = Utilities.getUuid() + Utilities.getUuid();
  cache.put('s_' + t, JSON.stringify({ k: u.key, s: sid }), 21600);
  var out = { ok: true, token: t, sid: sid, me: me_(u) };
  if (want && want.date && !u.must) { try { out.sum = getSummary(t, want); } catch (e) {} }          // the first summary comes with the login answer (one round trip less)
  return out;
}
function getMe(token) {
  var raw = token ? CacheService.getScriptCache().get('s_' + token) : null;
  if (raw && raw.indexOf('"dh"') > 0) return dhMe_(dhAuth_(token, true));
  return me_(auth_(token, true));
}
function logout(token) { CacheService.getScriptCache().remove('s_' + token); return true; }
function changePassword(token, oldPw, newPw) {
  var raw = token ? CacheService.getScriptCache().get('s_' + token) : null;
  if (raw && raw.indexOf('"dh"') > 0) {
    var h = dhAuth_(token, true);
    if (!chk_(h.hash, String(oldPw || ''))) throw new Error('The old password is wrong.');
    newPw = String(newPw || ''); if (newPw.length < 6) throw new Error('The new password must be at least 6 characters.');
    dhSetPass_(h, newPw, false); return true;
  }
  var u = auth_(token, true);
  if (!chk_(u.pass, String(oldPw || ''))) throw new Error('The old password is wrong.');
  newPw = String(newPw || '');
  if (newPw.length < 6) throw new Error('The new password must be at least 6 characters.');
  setPass_(u, newPw, false);
  return true;
}
function adminUsers(token) {
  var u = auth_(token); if (!u.admin) throw new Error('Admin only.');
  return users_().map(function (x) { return { name: x.name, display: x.display, role: x.role, scope: x.scope.join(','), canEdit: x.canEdit, active: x.active, admin: !!x.admin }; });
}
function adminReset(token, username) {
  var a = auth_(token); if (!a.admin) throw new Error('Admin only.');
  var u = users_().filter(function (x) { return x.key === String(username).trim().toLowerCase(); })[0];
  if (!u) throw new Error('User not found.');
  var t = 'Tx@' + rnd_(6);
  setPass_(u, t, true);          // the password must be changed at the next login
  return t;
}
function adminSetActive(token, username, active) {
  var a = auth_(token); if (!a.admin) throw new Error('Admin only.');
  var u = users_().filter(function (x) { return x.key === String(username).trim().toLowerCase(); })[0];
  if (!u) throw new Error('User not found.');
  if (u.key === a.key) throw new Error('You cannot deactivate your own ID.');
  ss_().getSheetByName(USERS).getRange(u.row, 6).setValue(!!active); usersBump_();
  return true;
}

function parseScope_(sc) {                              // "AZAMGARH,MAU" or ["AZAMGARH","MAU"] or "ALL" -> clean list of workshops
  var a = Array.isArray(sc) ? sc : String(sc || '').split(','), out = [], W = ws_();
  a.forEach(function (x) {
    var t = String(x || '').trim().toUpperCase(); if (!t) return;
    if (t === 'ALL') { out = ['ALL']; return; }
    var k = W.indexOf(t) >= 0 ? t : wsKey_(t); if (!k) throw new Error('Workshop "' + t + '" is not in the workshop list.');
    if (out.indexOf('ALL') < 0 && out.indexOf(k) < 0) out.push(k);
  });
  if (out.indexOf('ALL') >= 0) return ['ALL'];
  if (!out.length) throw new Error('Choose at least one workshop.');
  return out;
}
// additional charge: give one user the data of two or more workshops (not an Admin)
function adminSetScope(token, username, scope) {
  var a = auth_(token); if (!a.admin) throw new Error('Admin only.');
  var u = users_().filter(function (x) { return x.key === String(username).trim().toLowerCase(); })[0];
  if (!u) throw new Error('User not found.');
  if (u.admin) throw new Error('The Admin already sees every workshop.');
  var list = parseScope_(scope), l = lock_();
  try { ss_().getSheetByName(USERS).getRange(u.row, 7).setNumberFormat('@').setValue(list.join(',')); usersBump_(); bump_(); } finally { l.releaseLock(); }
  return list.join(',');
}
function adminAddUser(token, name, display, role, scope, canEdit) {
  var a = auth_(token); if (!a.admin) throw new Error('Admin only.');
  name = String(name || '').trim();
  if (!/^[A-Za-z0-9._@\-]{3,30}$/.test(name)) throw new Error('User ID must be 3-30 characters (A-Z, 0-9, . _ - @), no spaces.');
  var l = lock_();
  try {
    if (users_().some(function (x) { return x.key === name.toLowerCase(); })) throw new Error('This User ID already exists in your workshop list. Choose a different one.');
    scope = parseScope_(scope).join(',');
    var t = 'Tx@' + rnd_(6), sh = ss_().getSheetByName(USERS), r = Math.max(sh.getLastRow(), 1) + 1;
    sh.getRange(r, 1, 1, 9).setNumberFormat('@');
    sh.getRange(r, 1, 1, 9).setValues([[name, mk_(t), String(role || 'AE/JE').trim(), String(display || name).trim(), !!canEdit, true, scope, false, true]]);
    regAddUser_(name, SS_ID, String(role || 'AE/JE').trim()); usersBump_();
    return t;
  } finally { l.releaseLock(); }
}

/* ---------------- data read / write ---------------- */
function readAll_(keepBlank) {
  var sh = ss_().getSheetByName(DATA), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, NC).getValues().map(function (r) {
    return r.map(function (v, i) { return norm_(i, v); });
  }).filter(function (r) { return keepBlank || r[I_PR] !== ''; });
}
function writeRows_(sh, startRow, rows) {
  if (!rows.length) return;
  var need = startRow + rows.length - 1;
  if (need > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  sh.getRange(startRow, 1, rows.length, NC).setNumberFormat('@').setValues(rows);
}
function inScope_(u, wsk) { return u.scope.indexOf('ALL') >= 0 || (wsk && u.scope.indexOf(wsk) >= 0); }
function canEdit_(u, rowArr) { return u.canEdit && inScope_(u, wsKey_(rowArr[I_WS])); }

/* ---------------- tez index (summary / list ke liye) ----------------
 * Poori DATA sheet (36 columns x hazaaron rows) har click par padhna dheema tha. Ab sirf zaruri 8 cheezein
 * (workshop, division, capacity, PR NO, PR/Issue/Replacement/Return date) ek chhote "index" me rakhi jaati hain
 * aur CacheService me jama hoti hain. Koi bhi save/delete/edit hone par version badal jata hai, to index naya ban jata hai.
 * Index row = [workshop(0/1/2/-1), division#, capacity, PR NO, prd, iss, rep, ret]  (ret = RETURN DATE of the PR); position = sheet row - 2. Khali PR wali row = 0.   */
var IDX_MEM = null, IDX_CH = 1000, IDX_TTL = 900;
function ver_() {                                      // DATA version: changes after every save (summary / list caches use it)
  var c = CacheService.getScriptCache(), k = 'dv_' + SS_ID, v = c.get(k);
  if (!v) { v = Utilities.getUuid().slice(0, 8); c.put(k, v, 21600); }
  return v;
}
function ixv_() {                                      // INDEX generation: changes only when the whole index must be rebuilt (delete / paste / restore ...)
  var c = CacheService.getScriptCache(), k = 'iv_' + SS_ID, v = c.get(k);
  if (!v) { v = Utilities.getUuid().slice(0, 8); c.put(k, v, 21600); }
  return v;
}
function bump_() {                                     // big change: data version AND index are both thrown away (the next reader rebuilds the index once)
  IDX_MEM = null;
  try { var c = CacheService.getScriptCache(); c.put('dv_' + SS_ID, Utilities.getUuid().slice(0, 8), 21600); c.put('iv_' + SS_ID, Utilities.getUuid().slice(0, 8), 21600); } catch (e) {}
}
function bumpData_() {                                 // small in-place edit (saveRecord): only the data version changes, the index is patched row-wise (idxPatch_)
  try { var v = Utilities.getUuid().slice(0, 8); CacheService.getScriptCache().put('dv_' + SS_ID, v, 21600); if (IDX_MEM) IDX_MEM.dv = v; } catch (e) {}
}
function buildIdx_() {
  var sh = ss_().getSheetByName(DATA), n = sh ? sh.getLastRow() : 0, divs = [], dmap = {}, subs = [], smap = {}, rows = [], W = ws_();
  if (n >= 2) {
    var vals = sh.getRange(2, 1, n - 1, I_RET + 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      var r = vals[i], pr = norm_(I_PR, r[I_PR]);
      if (pr === '') { rows.push(0); continue; }
      var dv = divKey_(r[I_DIV]), di = dmap[dv];
      if (di === undefined) { di = divs.length; dmap[dv] = di; divs.push(dv); }
      var sb = divKey_(r[I_SUB]), si = smap[sb];
      if (si === undefined) { si = subs.length; smap[sb] = si; subs.push(sb); }
      rows.push([W.indexOf(wsKey_(r[I_WS])), di, capOf_(norm_(I_CAP, r[I_CAP])), pr,
        dk_(norm_(I_PRD, r[I_PRD])), dk_(norm_(I_ISS, r[I_ISS])), dk_(norm_(I_REP, r[I_REP])), dk_(norm_(I_RET, r[I_RET])), si, isDup_(r[I_FIRM]) ? 1 : 0]);
    }
  }
  return { divs: divs, subs: subs, rows: rows };
}
// Reads the index from the cache only (NEVER builds it). null = not in cache / does not match the sheet.
// "jd" key = small list of rows edited in place since the index was built (position -> index row); applied on top of the cached rows.
function idxLoad_() {
  // A warm server instance keeps the index in memory, but ANOTHER instance may have saved since then (it only bumps the cache version / writes a patch).
  // So the memory copy is used only while the cache versions are the same as when it was loaded; otherwise it is read again from the cache.
  var ivNow = '', dvNow = '';
  try { ivNow = ixv_(); dvNow = ver_(); } catch (e) {}
  if (IDX_MEM && IDX_MEM.iv === ivNow && IDX_MEM.dv === dvNow) return IDX_MEM;
  IDX_MEM = null;
  var cnt = -1, ix = null;
  try { var sh0 = ss_().getSheetByName(DATA); cnt = sh0 ? Math.max(sh0.getLastRow() - 1, 0) : 0; } catch (e) { return null; }
  try {
    var cc = CacheService.getScriptCache(), pre = 'jy' + SS_ID + ixv_(), m = cc.get(pre);
    if (!m) return null;
    var meta = JSON.parse(m), keys = [], k;
    for (k = 0; k < meta.n; k++) keys.push(pre + '_' + k);
    var got = cc.getAll(keys), rows = [];
    for (k = 0; k < meta.n; k++) { var s = got[keys[k]]; if (s === null || s === undefined) return null; rows = rows.concat(JSON.parse(s)); }
    if (meta.len !== cnt) return null;                // sheet row count changed (someone added / removed rows by hand)
    var dj = cc.get('jd' + SS_ID + ixv_());
    if (dj) {
      var o = JSON.parse(dj);
      Object.keys(o).forEach(function (p) {            // only when the row at that position is still the same PR
        var x = o[p], cur = rows[+p];
        if (cur && x && cur[0] === x[0] && cur[3] === x[3]) rows[+p] = x;
      });
    }
    ix = { divs: meta.divs, subs: meta.subs || [], rows: rows, iv: ivNow, dv: dvNow };
  } catch (e) { return null; }
  return (IDX_MEM = ix);
}
function idxStore_(ix, pre) {
  try {
    var cc = CacheService.getScriptCache(), nch = Math.ceil(ix.rows.length / IDX_CH), put = {}, cnt2 = 0, k;
    for (k = 0; k < nch; k++) {
      put[pre + '_' + k] = JSON.stringify(ix.rows.slice(k * IDX_CH, (k + 1) * IDX_CH));
      if (++cnt2 === 10) { cc.putAll(put, IDX_TTL); put = {}; cnt2 = 0; }
    }
    if (cnt2) cc.putAll(put, IDX_TTL);
    cc.put(pre, JSON.stringify({ n: nch, len: ix.rows.length, divs: ix.divs, subs: ix.subs }), IDX_TTL);     // meta last: an unfinished index is never used
  } catch (e) {}
}
// Single flight: when the cache is empty only ONE call reads the whole sheet, the others wait for it (before: every parallel call rebuilt it).
function idx_() {
  var ix = idxLoad_();
  if (ix) return ix;
  var cc = null, fk = '', iv = ixv_();
  try { cc = CacheService.getScriptCache(); fk = 'ib_' + SS_ID + iv; } catch (e) { cc = null; }
  if (cc && cc.get(fk)) {                               // another call is building it right now: wait (max ~20 s), then use its result
    for (var w = 0; w < 40 && !ix; w++) {
      Utilities.sleep(500);
      ix = idxLoad_();
      if (!ix && !cc.get(fk)) break;                    // the builder finished or failed
    }
    if (ix) return ix;
  }
  if (cc) { try { cc.put(fk, '1', 60); } catch (e) {} }
  try {
    ix = buildIdx_();
    if (ixv_() === iv) { idxStore_(ix, 'jy' + SS_ID + iv); try { cc.remove('jd' + SS_ID + iv); } catch (e) {} }     // a big change during the build -> do not store the stale copy; a fresh index needs no old in-place patches
  } finally { if (cc) { try { cc.remove(fk); } catch (e) {} } }
  ix.iv = iv; ix.dv = ver_();
  return (IDX_MEM = ix);
}
// 0-based position of the PR in the index (the sheet row = position + 2), -1 if not found
function idxFind_(ix, ws, pr) {
  var want = rk_(ws, pr), p = String(pr);
  for (var i = 0; i < ix.rows.length; i++) {
    var r = ix.rows[i];
    if (r && String(r[3]) === p && rk_(wsAt_(r[0]), r[3]) === want) return i;
  }
  return -1;
}
// After saveRecord: change only this row inside the cached index (no re-read of the sheet).
function idxPatch_(ix, pos, nv) {
  try {
    var cur = ix.rows[pos]; if (!cur) return false;
    var nr = cur.slice();
    nr[4] = dk_(nv[I_PRD]); nr[5] = dk_(nv[I_ISS]); nr[6] = dk_(nv[I_REP]); nr[7] = dk_(nv[I_RET]); nr[9] = isDup_(nv[I_FIRM]) ? 1 : 0;
    var cc = CacheService.getScriptCache(), key = 'jd' + SS_ID + ixv_(), o = {}, s = cc.get(key);
    if (s) o = JSON.parse(s);
    o[pos] = nr;
    if (Object.keys(o).length > 400) return false;       // too many in-place edits: rebuild the index in full
    cc.put(key, JSON.stringify(o), 21600);
    ix.rows[pos] = nr; IDX_MEM = ix;
    return true;
  } catch (e) { return false; }
}
// index ki row-positions (0-based) ki poori 36-column rows. Agar sheet index se mel na khaye (kisi ne row hata di) to null
function fetchRows_(ix, idxs) {
  var out = {};
  if (!idxs.length) return out;
  try {
    var sh = ss_().getSheetByName(DATA), s = idxs.slice().sort(function (a, b) { return a - b; }), runs = [], st = s[0], pv = s[0], want = {}, k;
    for (k = 0; k < s.length; k++) want[s[k]] = 1;
    for (k = 1; k < s.length; k++) { if (s[k] - pv > 300) { runs.push([st, pv]); st = s[k]; } pv = s[k]; }
    runs.push([st, pv]);
    if (runs.length > 8) runs = [[s[0], s[s.length - 1]]];          // bahut bikhri hui hain -> ek baar me padh lo
    for (var q = 0; q < runs.length; q++) {
      var a = runs[q][0], b = runs[q][1], vals = sh.getRange(a + 2, 1, b - a + 1, NC).getValues();
      for (var j = 0; j < vals.length; j++) {
        var i = a + j; if (!want[i]) continue;
        var row = vals[j].map(function (v, c) { return norm_(c, v); });
        if (row[I_PR] !== ix.rows[i][3]) return null;
        out[i] = row;
      }
    }
    return out;
  } catch (e) { return null; }
}
// sirf WORKSHOP aur PR NO columns (poori row nahi) - save / delete me position dhoondhne ke liye
function keys_(sh) {
  var n = sh.getLastRow();
  if (n < 2) return [];
  var a = sh.getRange(2, I_WS + 1, n - 1, 1).getValues(), p = sh.getRange(2, I_PR + 1, n - 1, 1).getValues();
  return a.map(function (r, i) { return [norm_(I_WS, r[0]), norm_(I_PR, p[i][0])]; });
}
function readRow_(sh, rowNum) {
  return sh.getRange(rowNum, 1, 1, NC).getValues()[0].map(function (v, i) { return norm_(i, v); });
}

function load_(u, c) {
  var ix = idx_(), out = [];
  for (var i = 0; i < ix.rows.length; i++) {
    var r = ix.rows[i]; if (!r) continue;
    var w = wsAt_(r[0]);
    if (!inScope_(u, w)) continue;
    out.push({ i: i, ws: w, div: ix.divs[r[1]], sub: (ix.subs || [])[r[8]] || '', cap: r[2], pr: r[3], prd: r[4], iss: r[5], rep: r[6], ret: r[7] || 0, dup: r[9] ? 1 : 0 });
  }
  // summary / list ki date 2 saal se purani ho to ARCHIVE ka data bhi apne aap jod do
  if (c && !c.noArc && Math.min(c.D, c.F) < arcLimit_()) archRead_(Math.min(c.D, c.F)).forEach(function (a) {
    var w = wsKey_(a[I_WS]);
    if (!inScope_(u, w)) return;
    out.push({ a: a, ws: w, div: divKey_(a[I_DIV]), sub: divKey_(a[I_SUB]), cap: capOf_(a[I_CAP]), pr: a[I_PR], prd: dk_(a[I_PRD]), iss: dk_(a[I_ISS]), rep: dk_(a[I_REP]), ret: dk_(a[I_RET]), dup: isDup_(a[I_FIRM]) ? 1 : 0, arch: true });
  });
  return out;
}

/* ---------------- paste (upsert by PR NO) ---------------- */
function mapCols_(rows) {
  for (var h = 0; h < Math.min(3, rows.length); h++) {
    var hdr = rows[h].map(function (c) { return String(c).replace(/\s+/g, ' ').trim().toUpperCase(); });
    if (hdr.indexOf('PR NO') >= 0 || hdr.indexOf('PR NO.') >= 0) {
      var idx = COLS.map(function (c) { return hdr.indexOf(c.toUpperCase()); });
      if (idx[I_PR] < 0) idx[I_PR] = Math.max(hdr.indexOf('PR NO'), hdr.indexOf('PR NO.'));
      if (idx[I_PRD] < 0) throw new Error('The "PR DATE" column was not found in the header.');
      return { start: h + 1, idx: idx };
    }
  }
  var off = (rows[0] && !wsKey_(rows[0][0]) && wsKey_(rows[0][1])) ? 1 : 0;   // pehla "#" column ho (doosre column me workshop ho) to chhod do
  return { start: 0, idx: COLS.map(function (c, i) { return i + off; }) };
}
function rk_(ws, pr) { return wsKey_(ws) + '|' + pr; }          // alag workshop ka same PR NO aapas me na takraye
function merge_(u, rows) {
  var m = mapCols_(rows), sh = ss_().getSheetByName(DATA);
  var all = readAll_(true), pos = {};
  all.forEach(function (r, i) { if (r[I_PR] !== '') pos[rk_(r[I_WS], r[I_PR])] = i; });
  var one = oneWs_(u);
  var added = [], updated = 0, same = 0, skipped = [], skipN = 0, rs = [];
  for (var k = m.start; k < rows.length; k++) {
    var src = rows[k];
    if (!src.some(function (v) { return String(v).trim() !== ''; })) continue;
    var rec = m.idx.map(function (j, i) { return j >= 0 ? norm_(i, src[j]) : ''; });
    dupFix_(rec);
    if (!rec[I_WS] && one) rec[I_WS] = one;                   // Excel me Workshop column na ho to apna workshop maan lo
    { var cw0 = wsKey_(rec[I_WS]); if (cw0) rec[I_WS] = cw0; }   // full name (e.g. "EWC Azamgarh") -> registered workshop name
    var why = '', foreign = false;
    if (!rec[I_PR]) why = 'PR NO is empty';
    else if (!dk_(rec[I_PRD])) why = 'PR DATE is invalid/empty (DD.MM.YYYY required)';
    else if (!wsKey_(rec[I_WS])) why = 'Workshop is empty/not recognised (AZAMGARH / MAU / BALLIA)';
    else if (!inScope_(u, wsKey_(rec[I_WS]))) foreign = true;           // a row of another workshop of the same Admin: saved there if the PR is new, skipped if it already exists
    if (why) { skipN++; if (skipped.length < 15) skipped.push('Line ' + (k + 1) + ' (PR ' + (rec[I_PR] || '-') + '): ' + why); continue; }
    var key = rk_(rec[I_WS], rec[I_PR]);
    if (foreign && (pos[key] !== undefined || archIdx_()[key])) { skipN++; if (skipped.length < 15) skipped.push('Line ' + (k + 1) + ' (PR ' + rec[I_PR] + '): this PR already exists in workshop ' + wsKey_(rec[I_WS]) + ' - skipped'); continue; }
    if (pos[key] === undefined) restore_(key, all, pos, added, rs);   // PR ARCHIVE me ho to wapas DATA me le aao
    if (pos[key] !== undefined) {
      var ex = all[pos[key]], ch = false;
      for (var i = 0; i < NC; i++) if (rec[i] !== '' && rec[i] !== ex[i]) { ex[i] = rec[i]; ch = true; }
      if (ch) updated++; else same++;
    } else {
      pos[key] = all.length; all.push(rec); added.push(rec);
    }
  }
  if (updated > 0) writeRows_(sh, 2, all);
  else if (added.length) writeRows_(sh, Math.max(sh.getLastRow(), 1) + 1, added);
  if (rs.length) archDel_(rs);
  bump_();
  return { added: added.length - rs.length, updated: updated, same: same, skipped: skipN, notes: skipped };
}
function saveRows(token, rows) {
  var u = auth_(token);
  if (!u.canEdit) throw new Error('You do not have permission to paste data.');
  if (!rows || !rows.length) throw new Error('Nothing was pasted.');
  var l = lock_();
  try { return merge_(u, rows); } finally { l.releaseLock(); }
}

/* ---------------- Excel-like data sheet (grid) ---------------- */
// Workshop login = sirf apne workshop ki rows, Admin (scope ALL) = sabki rows
// p = { date: 'DD.MM.YYYY', all: true/false } -> date server par hi filter hoti hai (PR DATE ke hisaab se),
// isliye sheet badi ho jaye tab bhi us din ka poora data aata hai. "All dates" me sabse nayi 6000 rows aati hain.
function getData(token, p) {
  var u = auth_(token), rows = [];
  p = p || {};
  var dn = (!p.all && dk_(p.date)) ? dk_(p.date) : 0;
  if (dn) {                                           // ek date: index se row dhoondo, sirf wahi rows padho (tez)
    for (var t = 0; t < 2; t++) {
      var ix = idx_(), hit = [];
      for (var i0 = 0; i0 < ix.rows.length; i0++) {
        var x = ix.rows[i0]; if (!x || x[4] !== dn) continue;
        if (!inScope_(u, wsAt_(x[0]))) continue;
        hit.push(i0);
      }
      hit.sort(function (a, b) { return b - a; });    // sheet me baad me judi row upar
      var take = hit.slice(0, 6000), got = fetchRows_(ix, take);
      if (got) return { cols: COLS, canEdit: u.canEdit, scope: u.scope, one: oneWs_(u), total: hit.length, rows: take.map(function (i) { return got[i]; }) };
      bump_();
    }
  }
  var d = dn ? String(p.date) : '';
  readAll_().forEach(function (a, i) {
    if (!inScope_(u, wsKey_(a[I_WS]))) return;
    if (d && a[I_PRD] !== d) return;
    rows.push({ a: a, k: dk_(a[I_PRD]), i: i });
  });
  // sabse nayi upar: pehle PR DATE (nayi se purani), same date me jo sheet me baad me judi wo upar
  rows.sort(function (x, y) { return (y.k - x.k) || (y.i - x.i); });
  rows = rows.map(function (o) { return o.a; });
  return { cols: COLS, canEdit: u.canEdit, scope: u.scope, one: oneWs_(u), total: rows.length, rows: rows.slice(0, 6000) };
}
function otherWs_(u, w) {
  return 'WARNING - this row is for workshop "' + wsKey_(w) + '", but you can only save entries of ' + u.scope.join(' / ') + '. It was NOT saved.';
}
function oneWs_(u) { return (u.scope.indexOf('ALL') < 0 && u.scope.length === 1 && ws_().indexOf(u.scope[0]) >= 0) ? u.scope[0] : ''; }
// items = [{k: purani "WORKSHOP|PRNO" (nayi row ho to ''), r: [36 values]}]  -> row poori replace hoti hai (khali cell = mita hua)
function saveGrid(token, items) {
  try { return saveGridRun_(token, items); }
  catch (e) {
    if (String((e && e.message) || e).indexOf('IDXSTALE') < 0) throw e;       // the cached index did not match the sheet (nothing was written yet): once more, reading the sheet itself
    IDX_MEM = null; return saveGridRun_(token, items, true);
  }
}
function saveGridRun_(token, items, noIx) {
  var u = auth_(token);
  if (!u.canEdit) throw new Error('You do not have permission to change data.');
  if (!items || !items.length) return { added: 0, updated: 0, skipped: 0, notes: [], keys: [], bad: [], rows: {} };
  var l = lock_(), fast = false;                                         // fast = sirf maujooda rows ke cells badle -> index poora fenkna nahi padega
  try {
    var sh = ss_().getSheetByName(DATA), all, pos = {}, one = oneWs_(u), snap = {};   // snap = badalne se pehle ki row (position -> purani row)
    // thode rows (roz ka autosave): sirf WORKSHOP + PR NO columns padho, jis row ko badalna hai wahi poori padho (tez).
    // Bahut rows (bada paste): pehle jaisa poori sheet padho.
    var lazy = items.length <= 15;
    var ixu = (lazy && !noIx) ? idxLoad_() : null, WSN = ws_();            // cached index: row positions come from it, the sheet's WORKSHOP + PR NO columns are NOT read at all
    if (lazy && ixu) {
      all = new Array(ixu.rows.length);
      ixu.rows.forEach(function (x, i) { if (x && x[3] !== '' && x[3] != null && WSN[x[0]]) pos[WSN[x[0]] + '|' + x[3]] = i; });
    } else if (lazy) {
      var ks = keys_(sh); all = new Array(ks.length);
      ks.forEach(function (k, i) { if (k[1] !== '') pos[rk_(k[0], k[1])] = i; });
    } else {
      all = readAll_(true);
      all.forEach(function (r, i) { if (r[I_PR] !== '') pos[rk_(r[I_WS], r[I_PR])] = i; });
    }
    function row_(i) {
      if (all[i] === undefined) {
        all[i] = readRow_(sh, i + 2);
        if (ixu) { var x = ixu.rows[i]; if (!x || wsKey_(all[i][I_WS]) !== WSN[x[0]] || all[i][I_PR] !== String(x[3])) throw new Error('IDXSTALE'); }
      }
      return all[i];
    }
    var added = [], upd = {}, skipN = 0, notes = [], keys = [], updN = 0, sameN = 0, rs = [], bad = [], fin = {};
    function skip(n, pr, why) { skipN++; bad.push(n); if (notes.length < 15) notes.push('Line ' + (n + 1) + ' (PR ' + (pr || '-') + '): ' + why); }
    items.forEach(function (it, n) {
      var src = it.r || [], rec = [], org = it.o && it.o.length ? it.o : null;
      for (var i = 0; i < NC; i++) rec.push(norm_(i, src[i]));
      dupFix_(rec);
      if (!rec[I_WS] && one) rec[I_WS] = one;
      { var cw1 = wsKey_(rec[I_WS]); if (cw1) rec[I_WS] = cw1; }
      var why = '', foreign = false;
      if (!rec[I_PR]) why = 'PR NO is empty';
      else if (!dk_(rec[I_PRD])) why = 'PR DATE is invalid/empty (DD.MM.YYYY required)';
      else if (!wsKey_(rec[I_WS])) why = 'Workshop is empty/not recognised (AZAMGARH / MAU / BALLIA)';
      else if (!inScope_(u, wsKey_(rec[I_WS]))) { if (it.k) why = otherWs_(u, rec[I_WS]); else foreign = true; }      // new row of another workshop of the same Admin: saved there only if the PR is new
      else for (var d = 0; d < DATE_IDX.length; d++) { var v = rec[DATE_IDX[d]]; if (v !== '' && !dk_(v)) { why = COLS[DATE_IDX[d]] + ' date must be written as DD.MM.YYYY'; break; } }
      if (!why) MOB_IDX.forEach(function (j) {                         // mobile: sirf wahi check jo is baar badla gaya
        if (why || rec[j] === '' || (org && rec[j] === norm_(j, org[j]))) return;
        var m = mobClean_(rec[j]);
        if (m === null) why = COLS[j] + ' must be 10 digits (E+ notation is not accepted)'; else rec[j] = m;
      });
      if (why) return skip(n, rec[I_PR], why);
      var key = rk_(rec[I_WS], rec[I_PR]), old = it.k ? pos[it.k] : undefined;
      if (foreign && (pos[key] !== undefined || archIdx_()[key])) return skip(n, rec[I_PR], 'this PR already exists in workshop ' + wsKey_(rec[I_WS]) + ' - skipped');
      if (!it.k && pos[key] === undefined) restore_(key, all, pos, added, rs);   // PR ARCHIVE me ho to wapas DATA me le aao
      if (it.k && old === undefined) return skip(n, rec[I_PR], 'the old entry was not found (did someone delete it?)');
      if (old !== undefined) {                                           // purani row edit
        if (!canEdit_(u, row_(old))) return skip(n, rec[I_PR], 'You do not have permission to change this entry');
        var mrg = rec;
        if (org) {                                                       // sirf wahi cell badlo jo user ne badle: purani (stale) row se kisi aur ka data na mite
          mrg = row_(old).slice();
          for (var j = 0; j < NC; j++) if (rec[j] !== norm_(j, org[j])) mrg[j] = rec[j];
          key = rk_(mrg[I_WS], mrg[I_PR]);
        }
        if (key !== it.k && pos[key] !== undefined) return skip(n, rec[I_PR], 'This PR NO already exists in this workshop');
        var chg = false, cur = row_(old); for (var q = 0; q < NC; q++) if (mrg[q] !== cur[q]) { chg = true; break; }
        if (!chg && key === it.k) { sameN++; keys.push(key); fin[n] = key; return; }
        if (!snap[old]) snap[old] = cur;
        all[old] = mrg; upd[old] = 1; updN++;
        if (key !== it.k) { delete pos[it.k]; pos[key] = old; }
      } else if (pos[key] !== undefined) {                              // nayi row, par PR NO pehle se hai -> update (khali cell purana data nahi mitata)
        var ex = row_(pos[key]), ch = false, exOld = ex.slice();
        if (!canEdit_(u, ex)) return skip(n, rec[I_PR], 'You do not have permission to change this entry');
        for (var j2 = 0; j2 < NC; j2++) if (rec[j2] !== '' && rec[j2] !== ex[j2]) { ex[j2] = rec[j2]; ch = true; }
        if (ch) { if (!snap[pos[key]]) snap[pos[key]] = exOld; upd[pos[key]] = 1; updN++; } else sameN++;
      } else { pos[key] = all.length; all.push(rec); added.push(rec); }
      keys.push(key); fin[n] = key;
    });
    var uk = Object.keys(upd).map(Number).sort(function (a, b) { return a - b; });
    if (uk.length > 40) writeRows_(sh, uk[0] + 2, all.slice(uk[0], uk[uk.length - 1] + 1));   // sirf badli hui rows ka hissa likho, poori sheet nahi
    else writeRuns_(sh, uk, all);                                                             // paas-paas ki rows ek saath
    if (added.length) writeRows_(sh, all.length - added.length + 2, added);
    if (rs.length) archDel_(rs);
    var rows = {}; Object.keys(fin).forEach(function (n) { var p = pos[fin[n]]; if (p !== undefined) rows[n] = row_(p); });   // sheet me jo asal me save hua
    // TEZ RASTA: koi nayi row / row hatana nahi, aur workshop-division-substation-capacity-PR NO nahi badle (sirf driver, firm, date, remark jaise cells)
    // -> index (cache) ki sirf wahi rows patch karo. Pehle har save ke baad poora index fenk diya jata tha, to agla "Loading data..." poori sheet dobara padhta tha (10-20 sec).
    if (!added.length && !rs.length && uk.length && uk.length <= 40) {
      var ixc = idxLoad_();                                                   // sirf cache se (kabhi build nahi karta)
      if (ixc) {
        fast = true;
        uk.forEach(function (p) {
          var o = snap[p], nw = all[p];
          if (!fast) return;
          if (!o || !nw || wsKey_(o[I_WS]) !== wsKey_(nw[I_WS]) || divKey_(o[I_DIV]) !== divKey_(nw[I_DIV]) || divKey_(o[I_SUB]) !== divKey_(nw[I_SUB]) ||
              capOf_(o[I_CAP]) !== capOf_(nw[I_CAP]) || o[I_PR] !== nw[I_PR]) fast = false;
        });
        if (fast) uk.forEach(function (p) { if (fast && !idxPatch_(ixc, p, all[p])) fast = false; });
      }
    }
    return { added: added.length - rs.length, updated: updN, same: sameN, skipped: skipN, notes: notes, keys: keys, bad: bad, rows: rows };
  } finally { if (fast) bumpData_(); else bump_(); l.releaseLock(); }
}
// sorted row-positions (0-based) ko paas-paas ke tukdon me jod kar likho (har row ke liye alag call nahi)
function writeRuns_(sh, idxs, all) {
  var st = 0;
  for (var k = 1; k <= idxs.length; k++) {
    if (k === idxs.length || idxs[k] !== idxs[k - 1] + 1) {
      writeRows_(sh, idxs[st] + 2, idxs.slice(st, k).map(function (i) { return all[i]; }));
      st = k;
    }
  }
}
function deleteRecord(token, ws, pr) {
  var u = auth_(token), l = lock_();
  try {
    var sh = ss_().getSheetByName(DATA), keys = keys_(sh), want = rk_(ws, String(pr).trim());
    for (var i = 0; i < keys.length; i++) {
      if (keys[i][1] !== '' && rk_(keys[i][0], keys[i][1]) === want) {
        if (!canEdit_(u, [keys[i][0]])) throw new Error('You do not have permission to delete this entry.');
        sh.deleteRow(i + 2); return true;
      }
    }
    throw new Error('PR ' + pr + ' not found.');
  } finally { bump_(); l.releaseLock(); }
}

/* ---------------- edit one record ---------------- */
function saveRecord(token, ws, pr, vals) {
  var u = auth_(token), l = lock_(), wrote = false, patched = false, ix = null, ri = -1;
  try {
    var sh = ss_().getSheetByName(DATA), row = null, want = rk_(ws, String(pr).trim());
    ix = idxLoad_();                                      // cache only: a save never builds the index
    if (ix) {                                             // fast: position from the index, then read ONLY that row to confirm it
      var p0 = idxFind_(ix, ws, String(pr).trim());
      if (p0 >= 0) { var r0 = readRow_(sh, p0 + 2); if (rk_(r0[I_WS], r0[I_PR]) === want) { ri = p0 + 2; row = r0; } }
    }
    if (ri < 0) {                                         // fallback (index not cached / does not match): read the two key columns
      var keys = keys_(sh);
      for (var i = 0; i < keys.length; i++) if (keys[i][1] !== '' && rk_(keys[i][0], keys[i][1]) === want) { ri = i + 2; break; }
      if (ri < 0) throw new Error('PR ' + pr + ' not found.');
      row = readRow_(sh, ri);
    }
    if (!canEdit_(u, row)) throw new Error('You do not have permission to change this entry.');
    var nv = row.slice();
    Object.keys(vals || {}).forEach(function (k) {
      var i = +k; if (i < EDIT_FROM || i >= NC) return;
      var v = norm_(i, vals[k]);
      if (DATE_IDX.indexOf(i) >= 0 && v !== '' && !dk_(v)) throw new Error(COLS[i] + ' date must be written as DD.MM.YYYY.');
      if (MOB_IDX.indexOf(i) >= 0 && v !== '' && v !== row[i]) { var mc = mobClean_(v); if (mc === null) throw new Error(COLS[i] + ' must be 10 digits (E+ notation is not accepted).'); v = mc; }
      nv[i] = v;
    });
    dupFix_(nv);
    wrote = true;
    sh.getRange(ri, 1, 1, NC).setNumberFormat('@').setValues([nv]);
    if (ix) patched = idxPatch_(ix, ri - 2, nv);          // only this row of the cached index changes
    return nv.concat([1]);
  } finally {
    if (wrote) { if (patched) bumpData_(); else bump_(); }      // patched: only the summary / list caches are renewed, the index stays
    l.releaseLock();
  }
}

/* ---------------- summary ---------------- */
function ctx_(p) {
  var c = { D: dk_(p.date), F: dk_(p.from), T: dk_(p.to) };
  if (!c.D || !c.F || !c.T) throw new Error('Enter the date in DD.MM.YYYY format.');
  if (c.F > c.T) throw new Error('The "From" date cannot be later than the "To" date.');
  return c;
}
// 14 metrics. A PR that was RETURNED (RETURN DATE) is closed: it is no longer pending for issue / replacement.
var NM = 15;
function flags_(r, c) {
  var pr = r.prd, is = r.iss, rp = r.rep, rt = r.ret || 0, D = c.D, F = c.F, T = c.T;
  var open = function (x) { return rt === 0 || rt > x; };   // not returned as on date x
  if (r.dup) { var z = zero_(NM); z[14] = (pr >= F && pr <= T) ? 1 : 0; return z; }     // wrong / duplicate PR: only its own column
  // PR perspective: issue / replacement / return are counted for the PRs of the period PLUS the PRs that were already pending at the start (last month pendency).
  // So  LAST MONTH PENDENCY + TOTAL PR - TOTAL ISSUE = TOTAL PENDING FOR ISSUE, and when there is no last month pendency the issue can never be more than the PRs.
  var inB = (pr >= F && pr <= T) || (pr < F && (is === 0 || is >= F) && open(F - 1));
  return [
    pr < F && (is === 0 || is >= F) && open(F - 1),   // 0 LAST MONTH (opening) PENDENCY
    pr === D,                                   // 1 TODAY PR
    pr === D && is > 0,                         // 2 ISSUE AGAINST TODAY'S PR
    pr === D && rp > 0,                         // 3 REPLACEMENT AGAINST TODAY'S PR
    is === D,                                   // 4 TODAY ISSUE
    rp === D,                                   // 5 TODAY REPLACEMENT
    pr >= F && pr <= T,                         // 6 TOTAL PR
    inB && is > 0 && is <= T,                   // 7 TOTAL ISSUE        (PR perspective: of the PRs of the period + the opening pendency, how many are issued)
    inB && rp > 0 && rp <= T,                   // 8 TOTAL REPLACEMENT  (PR perspective)
    pr <= T && (is === 0 || is > T) && open(T),                     // 9 TOTAL PENDING FOR ISSUE
    is > 0 && is <= T && (rp === 0 || rp > T) && open(T),           // 10 TOTAL PENDING FOR REPLACEMENT
    rt === D,                                   // 11 TODAY RETURN
    inB && rt > 0 && rt <= T,                   // 12 TOTAL RETURN      (PR perspective)
    rp > 0 && rp <= T && (rt === 0 || rt > T),  // 13 TOTAL PENDING FOR RETURN (replaced, faulty transformer not returned yet)
    false                                       // 14 WRONG / DUPLICATE PR (set above)
  ];
}
function zero_(n) { var a = []; for (var i = 0; i < n; i++) a.push(0); return a; }
function addTo_(a, b) { for (var i = 0; i < a.length; i++) a[i] += b[i]; return a; }
function dayCount_(c) {
  function d(k) { return new Date(Math.floor(k / 10000), Math.floor(k / 100) % 100 - 1, k % 100); }
  return Math.round((d(c.T) - d(c.F)) / 86400000) + 1;
}

function getSummary(token, p) {
  var u = auth_(token), cc = CacheService.getScriptCache(), key = 'sm3_' + SS_ID + '_' + ver_() + '_' + u.scope.join(',').slice(0, 60) + '_' + [p.date, p.from, p.to].join('-'), hit = null;
  try { hit = bigGet_(cc, key); } catch (e) {}
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  var out = summaryCalc_(u, p);
  try { bigPut_(cc, key, JSON.stringify(out), 1800); } catch (e) {}
  return out;
}
function summaryCalc_(u, p) {
  var c = ctx_(p), recs = load_(u, c), WS = ws_(), nW = WS.length;
  var capSet = {}; STD_CAPS.forEach(function (k) { capSet[k] = 1; });
  recs.forEach(function (r) { capSet[r.cap] = 1; });
  var caps = Object.keys(capSet).map(Number).sort(function (a, b) { return (a || 1e9) - (b || 1e9); });

  var wsShown = WS.filter(function (w) { return inScope_(u, w); });
  var top = {}; WS.forEach(function (w) { top[w] = zero_(NM); });
  var divMap = {}, capI = {}, capR = {}, capP = {}, capT = {}, capRP = {}, divCap = {}, divWs = {}, pdw = {};      // pdw: workshop -> age bucket -> capacity -> pending for issue
  caps.forEach(function (k) { capI[k] = zero_(nW); capR[k] = zero_(nW); capP[k] = zero_(nW); capT[k] = zero_(nW); capRP[k] = zero_(nW); });

  recs.forEach(function (r) {
    var f = flags_(r, c), w = WS.indexOf(r.ws);
    var dv = divMap[r.div] || (divMap[r.div] = zero_(NM));
    var dc = divCap[r.div] || (divCap[r.div] = {});
    if (w >= 0 && !divWs[r.div]) divWs[r.div] = r.ws;                  // ye division kis workshop ki hai
    for (var i = 0; i < NM; i++) if (f[i]) { dv[i]++; if (w >= 0) top[r.ws][i]++; }
    if (w >= 0) { if (f[9]) capI[r.cap][w]++; if (f[10]) capR[r.cap][w]++; if (f[6]) capP[r.cap][w]++; if (f[12]) capT[r.cap][w]++; if (f[13]) capRP[r.cap][w]++; }
    if (f[6]) dc[r.cap] = (dc[r.cap] || 0) + 1;
    if (f[9] && w >= 0) { var ag = dayNum_(c.T) - dayNum_(r.prd), bi = ag <= 0 ? 0 : ag <= 2 ? 1 : ag <= 4 ? 2 : 3, pw = pdw[r.ws] || (pdw[r.ws] = [{}, {}, {}, {}]); pw[bi][r.cap] = (pw[bi][r.cap] || 0) + 1; }
  });

  function sumV(keys) { var s = zero_(NM); keys.forEach(function (k) { if (divMap[k]) addTo_(s, divMap[k]); }); return s; }
  function sumC(keys) { var s = {}; keys.forEach(function (k) { var d = divCap[k]; if (d) for (var cp in d) s[cp] = (s[cp] || 0) + d[cp]; }); return s; }

  // layout: har workshop ke neeche uski divisions (data se apne aap), phir workshop ka total
  var rowsD = [], rowsC = [], assigned = {}, allDivs = Object.keys(divMap).filter(Boolean);
  function natural(a, b) { return a.localeCompare(b, undefined, { numeric: true }); }
  function push(t, label, keys) { rowsD.push({ t: t, label: label, keys: keys, v: sumV(keys) }); rowsC.push({ t: t, label: label, keys: keys, c: sumC(keys) }); }
  wsShown.forEach(function (w) {
    var keys = allDivs.filter(function (d) { return divWs[d] === w; }).sort(natural);
    keys.forEach(function (d) { push('d', d, [d]); assigned[d] = 1; });
    push('w', w, keys);
  });
  var others = allDivs.filter(function (k) { return !assigned[k]; });
  var othersAny = others.some(function (k) { return divMap[k].some(function (x) { return x; }) || Object.keys(divCap[k] || {}).length; });
  if (othersAny) push('o', 'OTHER / NOT IN LIST', others);
  var allKeys = Object.keys(divMap);

  var topRows = wsShown.map(function (w) { return { label: w, ws: w, v: top[w] }; });
  var topTot = zero_(NM); topRows.forEach(function (r) { addTo_(topTot, r.v); });

  return {
    date: p.date, from: p.from, to: p.to, days: dayCount_(c), ws: wsShown, wsAll: WS, caps: caps,
    top: { rows: topRows, total: topTot },
    div: { rows: rowsD, total: sumV(allKeys) },
    divCap: { rows: rowsC, total: sumC(allKeys) },
    capIssue: capI, capRep: capR, capPR: capP, capRet: capT, capPendRet: capRP,
    pend: wsShown.map(function (w) {              // pending for issue, AGEING (PR date vs the To date) x capacity, per workshop
      var pw = pdw[w] || [{}, {}, {}, {}], B = [['0-24 hour', 0, 0], ['1-2 days', 1, 2], ['2-4 days', 3, 4], ['More than 4 days', 5, 1e9]];
      function key(n) { var d = new Date(n * 86400000); return d.getUTCFullYear() * 10000 + (d.getUTCMonth() + 1) * 100 + d.getUTCDate(); }
      return { ws: w, rows: B.map(function (b, i) { return { k: i, label: b[0], ph: key(dayNum_(c.T) - b[1]), pl: b[2] >= 1e9 ? 0 : key(dayNum_(c.T) - b[2]), v: caps.map(function (cp) { return pw[i][cp] || 0; }) }; }) };
    })
  };
}

/* ---------------- comparison (this period vs the same dates of the previous year) ---------------- */
function shiftYear_(k, n) {
  var y = Math.floor(k / 10000) + n, m = Math.floor(k / 100) % 100, d = k % 100;
  if (m === 2 && d === 29 && !((y % 4 === 0 && y % 100 !== 0) || y % 400 === 0)) d = 28;
  return y * 10000 + m * 100 + d;
}
function dayNum_(k) { return Date.UTC(Math.floor(k / 10000), Math.floor(k / 100) % 100 - 1, k % 100) / 86400000; }
function dmy_(k) { return p2_(k % 100) + '.' + p2_(Math.floor(k / 100) % 100) + '.' + Math.floor(k / 10000); }
function addMap_(a, b) { for (var k in b) a[k] = (a[k] || 0) + b[k]; return a; }
// mode: div (division wise) | sub (substation wise) | ws (workshop wise) | day (day wise). Counts PRs (PR DATE) capacity wise.
function getCompare(token, p) {
  var u = auth_(token), F = dk_(p.from), T = dk_(p.to), mode = String(p.mode || 'div');
  if (!F || !T) throw new Error('Enter the From / To dates in DD.MM.YYYY format.');
  if (F > T) throw new Error('The "From" date cannot be later than the "To" date.');
  if (['div', 'sub', 'ws', 'day'].indexOf(mode) < 0) throw new Error('Unknown comparison type.');
  var PF = shiftYear_(F, -1), PT = shiftYear_(T, -1), nd = dayNum_(T) - dayNum_(F) + 1;
  if (mode === 'day' && nd > 400) throw new Error('Day wise comparison: please choose a range of up to 400 days.');
  var useSt = statsReady_(), recs = load_(u, { D: PF, F: PF, T: T, noArc: useSt }), only = p.div ? divKey_(p.div) : '';
  if (useSt) statsIter_(PF, T, function (w, dv, sb, cp, d, n) { if (inScope_(u, w)) recs.push({ ws: w, div: dv, sub: sb, cap: cp, prd: d, dup: 0, n: n }); });
  var WS = ws_(), M = {}, order = [], divWs = {};
  function slot(key, o) { var s = M[key]; if (!s) { s = M[key] = { key: key, c: {}, p: {} }; for (var x in o) s[x] = o[x]; order.push(key); } return s; }
  if (mode === 'day') for (var d = 0; d < nd; d++) slot('d' + d, { off: d, w: {} });
  recs.forEach(function (r) {
    if (!r.prd || r.dup || (only && r.div !== only)) return;
    var cur = r.prd >= F && r.prd <= T, prv = r.prd >= PF && r.prd <= PT;
    if (!cur && !prv) return;
    if (r.ws && !divWs[r.div]) divWs[r.div] = r.ws;
    var s, nn = r.n || 1;
    if (mode === 'div') s = slot((r.ws || '') + '\u0001' + r.div, { div: r.div, ws: r.ws || '' });          // one line per (workshop, division)
    else if (mode === 'sub') s = slot(r.div + '\u0001' + r.sub, { div: r.div, sub: r.sub });
    else if (mode === 'ws') s = slot(r.ws || '-', { ws: r.ws || '-' });
    else {
      var sc = cur && M['d' + (dayNum_(r.prd) - dayNum_(F))], sp = prv && M['d' + (dayNum_(r.prd) - dayNum_(PF))];
      var wk = r.ws || '-';
      if (sc) { sc.c[r.cap] = (sc.c[r.cap] || 0) + nn; (sc.w[wk] || (sc.w[wk] = [0, 0]))[0] += nn; }
      if (sp) { sp.p[r.cap] = (sp.p[r.cap] || 0) + nn; (sp.w[wk] || (sp.w[wk] = [0, 0]))[1] += nn; }
      return;
    }
    if (cur) s.c[r.cap] = (s.c[r.cap] || 0) + nn;
    if (prv) s.p[r.cap] = (s.p[r.cap] || 0) + nn;
  });
  function natural(a, b) { return String(a).localeCompare(String(b), undefined, { numeric: true }); }
  var rows = [], tot = { c: {}, p: {} };
  if (mode === 'day') {
    order.forEach(function (k) { var s = M[k], dc = Math.round(dayNum_(F)) + s.off, base = new Date(dc * 86400000), pb = new Date((dayNum_(PF) + s.off) * 86400000);
      var kc = base.getUTCFullYear() * 10000 + (base.getUTCMonth() + 1) * 100 + base.getUTCDate(), kp = pb.getUTCFullYear() * 10000 + (pb.getUTCMonth() + 1) * 100 + pb.getUTCDate();
      rows.push({ t: 'd', label: dmy_(kc), plabel: dmy_(kp), from: dmy_(kc), to: dmy_(kc), pfrom: dmy_(kp), pto: dmy_(kp), c: s.c, p: s.p, wc: s.w }); addMap_(tot.c, s.c); addMap_(tot.p, s.p); });
  } else if (mode === 'ws') {
    order.sort(function (a, b) { var x = WS.indexOf(a), y = WS.indexOf(b); return (x < 0 ? 99 : x) - (y < 0 ? 99 : y); });
    order.forEach(function (k) { var s = M[k]; rows.push({ t: 'd', label: s.ws, ws: s.ws, c: s.c, p: s.p }); addMap_(tot.c, s.c); addMap_(tot.p, s.p); });
  } else if (mode === 'div') {
    var groups = {}, gl = [];
    order.forEach(function (k) { var s = M[k], w = s.ws || '', g = groups[w]; if (!g) { g = groups[w] = []; gl.push(w); } g.push(s); });
    gl.sort(function (a, b) { var x = WS.indexOf(a), y = WS.indexOf(b); return (x < 0 ? 99 : x) - (y < 0 ? 99 : y); });
    gl.forEach(function (w) {
      var g = groups[w].sort(function (a, b) { return natural(a.div, b.div); }), sc = {}, sp = {}, keys = [];
      g.forEach(function (s) { rows.push({ t: 'd', label: s.div, divs: [s.div], ws: w, c: s.c, p: s.p }); addMap_(sc, s.c); addMap_(sp, s.p); keys.push(s.div); });
      if (w) rows.push({ t: 'w', label: w + ' TOTAL', divs: keys, ws: w, c: sc, p: sp });          // sub-total of every workshop, right below its divisions
      addMap_(tot.c, sc); addMap_(tot.p, sp);
    });
  } else {                                                               // sub: division name on the left, then the substations of that division
    var byDiv = {}, dl = [];
    order.forEach(function (k) { var s = M[k], g = byDiv[s.div]; if (!g) { g = byDiv[s.div] = []; dl.push(s.div); } g.push(s); });
    dl.sort(natural);
    dl.forEach(function (dv) {
      var g = byDiv[dv].sort(function (a, b) { return natural(a.sub, b.sub); }), sc = {}, sp = {};
      g.forEach(function (s) { rows.push({ t: 'd', div: dv, label: s.sub || '(blank)', divs: [dv], subs: [s.sub], c: s.c, p: s.p }); addMap_(sc, s.c); addMap_(sp, s.p); });
      if (dl.length > 1) rows.push({ t: 'w', div: dv, label: dv + ' TOTAL', divs: [dv], c: sc, p: sp });
      addMap_(tot.c, sc); addMap_(tot.p, sp);
    });
  }
  var divs = {}; recs.forEach(function (r) { if (r.div) divs[r.div] = 1; });
  return { mode: mode, from: p.from, to: p.to, pfrom: dmy_(PF), pto: dmy_(PT), days: nd, wsList: WS.filter(function (w) { return inScope_(u, w); }), rows: rows, total: tot, divList: Object.keys(divs).sort(natural), archived: Math.min(F, PF) < cutoff_() };
}

// DISCOM head: day wise damage (PR) of every division, this year vs the same dates of the previous year
function getDiscomCompare(token, p) {
  var h = dhAuth_(token, false, 'comparison'), F = dk_(p.from), T = dk_(p.to);
  if (!F || !T) throw new Error('Enter the From / To dates in DD.MM.YYYY format.');
  if (F > T) throw new Error('The "From" date cannot be later than the "To" date.');
  var nd = dayNum_(T) - dayNum_(F) + 1;
  if (nd > 400) throw new Error('Please choose a range of up to 400 days.');
  var cc = CacheService.getScriptCache(), key = 'dc_' + h.discom.replace(/[^A-Z0-9]/g, '').slice(0, 30) + '_' + p.from + '-' + p.to, hit = null;
  try { hit = bigGet_(cc, key); } catch (e) {}
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  var PF = shiftYear_(F, -1), PT = shiftYear_(T, -1), seen = {}, names = [], cols = [];
  regAdmins_().forEach(function (r) {
    var sid = String(r[8]).trim(); if (!sid || seen[sid] || dnorm_(r[3]) !== h.discom) return;
    seen[sid] = 1;
    var name = String(r[9] || '').trim().toUpperCase() || ('(division name not set) ' + (String(r[4] || r[2]).trim() || sid.slice(0, 6)));
    var cur = [], prv = [], i; for (i = 0; i < nd; i++) { cur.push(0); prv.push(0); }
    function addD(d, n) {
      if (d >= F && d <= T) cur[dayNum_(d) - dayNum_(F)] += n;
      if (d >= PF && d <= PT) { var o = dayNum_(d) - dayNum_(PF); if (o >= 0 && o < nd) prv[o] += n; }
    }
    var st = false; try { useSheet_(sid); st = statsReady_(); } catch (e) {}
    (st ? sheetRows_(sid) : rowsWithArc_(sid, PF)).forEach(function (x) {      // older finished records live in ARCHIVE (monthly STATS after "Convert")
      if (!x[I_PR] || isDup_(x[I_FIRM])) return;
      var d = dk_(x[I_PRD]); if (d) addD(d, 1);
    });
    if (st) statsIter_(PF, T, function (w, dv, sb, cp, d, n) { addD(d, n); });
    cols.push({ name: name, cur: cur, prv: prv });
  });
  cols.sort(function (a, b) { return a.name.localeCompare(b.name, undefined, { numeric: true }); });
  var rows = [];
  for (var d = 0; d < nd; d++) {
    var bc = new Date((dayNum_(F) + d) * 86400000), bp = new Date((dayNum_(PF) + d) * 86400000);
    var kc = bc.getUTCFullYear() * 10000 + (bc.getUTCMonth() + 1) * 100 + bc.getUTCDate(), kp = bp.getUTCFullYear() * 10000 + (bp.getUTCMonth() + 1) * 100 + bp.getUTCDate();
    rows.push({ label: dmy_(kc), plabel: dmy_(kp), v: cols.map(function (x) { return [x.cur[d], x.prv[d]]; }) });
  }
  var out = { discom: h.discom, from: p.from, to: p.to, pfrom: dmy_(PF), pto: dmy_(PT), days: nd, divs: cols.map(function (x) { return x.name; }), rows: rows };
  try { bigPut_(cc, key, JSON.stringify(out), 120); } catch (e) {}
  return out;
}

/* ---------------- Kitty (stock of transformers) ----------------
 * Every workshop fills ITS OWN kitty day by day; the Division Head (Admin) only views and approves.
 *   KittyDay : Date | Workshop | Capacity | Issue | Repair | Issue to other workshop | Receive from other workshop | Previous year kitty | Updated on | Updated by
 *   KittyOpen: Date | Workshop | Capacity | Opening kitty | Set by | Set on      (an opening kitty that was set / approved on a date)
 *   KittyReq : Id | Requested on | Workshop | Date | By | Values(JSON) | Status | Decided by | Decided on
 * Balance kitty = Opening + Repair - Issue - Issue to other workshop + Receive from other workshop.
 * Opening kitty of a day = Balance of the previous day (carried forward automatically); it can only be changed by an approved request. */
var KD = 'KittyDay', KO = 'KittyOpen', KR = 'KittyReq', KP = 'KittyPrev';
var KP_HDR = ['Date', 'Workshop', 'Capacity', 'Balance kitty (last year)', 'Set by', 'Set on'];
var KD_HDR = ['Date', 'Workshop', 'Capacity', 'Issue', 'Repair', 'Issue to other workshop', 'Receive from other workshop', 'Previous year kitty', 'Updated on', 'Updated by'];
var KO_HDR = ['Date', 'Workshop', 'Capacity', 'Opening kitty', 'Set by', 'Set on'];
var KR_HDR = ['Id', 'Requested on', 'Workshop', 'Date', 'By', 'Values', 'Status', 'Decided by', 'Decided on'];
var KIT_CAPS = [10, 16, 25, 63, 100, 160, 250, 400, 630, 800, 1000];
function kitTab_(name, hdr) {
  var ss = ss_(), sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.getRange(1, 1, 1, hdr.length).setValues([hdr]).setFontWeight('bold').setBackground('#e3e9fb'); sh.setFrozenRows(1); sh.getRange(1, 1, sh.getMaxRows(), hdr.length).setNumberFormat('@'); }
  return sh;
}
function kitRows_(name, hdr) { var sh = kitTab_(name, hdr), n = sh.getLastRow(); return n >= 2 ? sh.getRange(2, 1, n - 1, hdr.length).getValues() : []; }
function kitWrite_(name, hdr, rows) {
  var sh = kitTab_(name, hdr);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, hdr.length).clearContent();
  if (!rows.length) return;
  if (sh.getMaxRows() < rows.length + 1) sh.insertRowsAfter(sh.getMaxRows(), rows.length + 1 - sh.getMaxRows());
  sh.getRange(2, 1, rows.length, hdr.length).setNumberFormat('@').setValues(rows);
}
/* KittyDays: ONE DAY = ONE CELL (all workshops and capacities of that date).  Date (YYYYMMDD) | {"WORKSHOP": {on, by, c: [[cap, issue, repair, iso, rfo], ...]}}
 * The first time it is needed the old KittyDay tab is copied into it (the old tab itself is left untouched). */
var KDS = 'KittyDays', KDS_HDR = ['Date (YYYYMMDD)', 'One day = one cell (all workshops, all capacities)'];
function kdsTab_() {
  var ss = ss_(), sh = ss.getSheetByName(KDS); if (sh) return sh;
  sh = kitTab_(KDS, KDS_HDR);
  var old = ss.getSheetByName(KD), n = old ? old.getLastRow() : 0;
  if (n >= 2) {
    var days = {};
    old.getRange(2, 1, n - 1, KD_HDR.length).getValues().forEach(function (r) {
      var w = wsKey_(r[1]), k = dk_(r[0]); if (!w || !k) return;
      var d = days[k] || (days[k] = {}), e = d[w] || (d[w] = { on: String(r[8] || ''), by: String(r[9] || ''), c: [] });
      e.c.push([+r[2], kitNum_(r[3]), kitNum_(r[4]), kitNum_(r[5]), kitNum_(r[6])]);
    });
    var rows = Object.keys(days).sort().map(function (k) { return [k, JSON.stringify(days[k])]; });
    if (rows.length) {
      if (sh.getMaxRows() < rows.length + 1) sh.insertRowsAfter(sh.getMaxRows(), rows.length + 1 - sh.getMaxRows());
      sh.getRange(2, 1, rows.length, 2).setNumberFormat('@').setValues(rows);
    }
  }
  return sh;
}
function kdsRows_() {                                  // the same row shape the old KittyDay tab had
  var sh = kdsTab_(), n = sh.getLastRow(), out = [];
  if (n < 2) return out;
  sh.getRange(2, 1, n - 1, 2).getValues().forEach(function (r) {
    var k = +r[0]; if (!k) return; var d; try { d = JSON.parse(r[1]); } catch (e) { return; }
    var ds = kitDs_(k);
    for (var w in d) (d[w].c || []).forEach(function (c) { out.push([ds, w, c[0], c[1], c[2], c[3], c[4], 0, d[w].on || '', d[w].by || '']); });
  });
  return out;
}
function kdsSetDay_(k, w, c, on, by) {                 // writes ONE cell
  var sh = kdsTab_(), n = sh.getLastRow(), row = 0, d = {};
  if (n >= 2) { var a = sh.getRange(2, 1, n - 1, 1).getValues(); for (var i = 0; i < a.length; i++) if (+a[i][0] === k) { row = i + 2; break; } }
  if (row) { try { d = JSON.parse(sh.getRange(row, 2).getValue()) || {}; } catch (e) { d = {}; } }
  d[w] = { on: on, by: by, c: c };
  if (!row) { row = Math.max(n, 1) + 1; if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), row - sh.getMaxRows()); sh.getRange(row, 1).setNumberFormat('@').setValue(String(k)); }
  sh.getRange(row, 2).setNumberFormat('@').setValue(JSON.stringify(d));
}
function kdsMap_(fn) {                                 // workshop renamed / deleted: change every day (only if the tab exists)
  var sh = ss_().getSheetByName(KDS), n = sh ? sh.getLastRow() : 0; if (n < 2) return;
  var rg = sh.getRange(2, 2, n - 1, 1), v = rg.getValues(), ch = 0;
  v.forEach(function (r) { var d; try { d = JSON.parse(r[0]); } catch (e) { return; } if (fn(d)) { r[0] = JSON.stringify(d); ch++; } });
  if (ch) rg.setNumberFormat('@').setValues(v);
}
function kitNum_(v) { var n = Math.round(+String(v == null ? '' : v).replace(/[^\d.\-]/g, '')); return isFinite(n) && n > 0 ? n : 0; }
function kitNow_() { return Utilities.formatDate(new Date(), TZ, 'dd.MM.yyyy HH:mm'); }
function kitToday_() { return Utilities.formatDate(new Date(), TZ, 'dd.MM.yyyy'); }
function kitDate_(d) {
  d = String(d || '').trim() || kitToday_();
  var k = dk_(d); if (!k) throw new Error('Enter a valid date (DD.MM.YYYY).');
  if (k > dk_(kitToday_())) throw new Error('Kitty cannot be filled for a future date.');
  return { s: d, k: k };
}
// per (workshop|capacity): opening adjustments, day movements, saved day figures
function kitCalc_(wsl) {
  var day = kdsRows_(), op = kitRows_(KO, KO_HDR), res = {}, ws = {}, has = {};
  wsl.forEach(function (w) { ws[w] = 1; });
  function slot(w, c) { var id = w + '|' + c; return res[id] || (res[id] = { adj: [], mv: [], day: {} }); }
  op.forEach(function (r) { var w = wsKey_(r[1]); if (!w || !ws[w]) return; var kk = dk_(r[0]); if (kk) slot(w, +r[2]).adj.push([kk, kitNum_(r[3])]); });
  day.forEach(function (r) {
    var w = wsKey_(r[1]); if (!w || !ws[w]) return;
    var kk = dk_(r[0]); if (!kk) return; has[w + '|' + kk] = 1;
    var x = slot(w, +r[2]), net = kitNum_(r[4]) - kitNum_(r[3]) - kitNum_(r[5]) + kitNum_(r[6]);
    x.mv.push([kk, net]); x.day[kk] = { issue: kitNum_(r[3]), repair: kitNum_(r[4]), iso: kitNum_(r[5]), rfo: kitNum_(r[6]), on: String(r[8] || ''), by: String(r[9] || '') };
  });
  return { slots: res, has: has };
}
function slotOpen_(x, k) {                         // opening kitty on date key k
  var ak = 0, av = 0, o;
  if (x) x.adj.forEach(function (a) { if (a[0] <= k && a[0] >= ak) { ak = a[0]; av = a[1]; } });
  o = av; if (x) x.mv.forEach(function (m) { if (m[0] >= ak && m[0] < k) o += m[1]; });
  return o;
}
function slotBal_(x, k) { var o = slotOpen_(x, k); if (x) x.mv.forEach(function (m) { if (m[0] === k) o += m[1]; }); return o; }
function kitDs_(k) { var y = Math.floor(k / 10000), m = Math.floor(k / 100) % 100, d = k % 100; return ('0' + d).slice(-2) + '.' + ('0' + m).slice(-2) + '.' + y; }
function kitKey_(dt) { return dt.getFullYear() * 10000 + (dt.getMonth() + 1) * 100 + dt.getDate(); }
function kitLy_(k) {                               // same date one year earlier (29 Feb -> 28 Feb)
  var y = Math.floor(k / 10000), m = Math.floor(k / 100) % 100, d = k % 100, t = new Date(y - 1, m - 1, d);
  if (t.getMonth() !== m - 1) t = new Date(y - 1, m - 1, 28);
  return t;
}
// previous year kitty = balance of the same date last year; if that date is not filled, the nearest earlier filled date (max 5 days back).
// Source 1: the workshop's own saved kitty of that date. Source 2: the figure the Division Head typed for last year (KittyPrev).
function kitPrevMap_(k, wsl, calc) {
  var kp = kitRows_(KP, KP_HDR), man = {}, out = {}, info = {}, base = kitLy_(k);
  kp.forEach(function (r) { var w = wsKey_(r[1]); if (w) man[w + '|' + dk_(r[0]) + '|' + (+r[2])] = kitNum_(r[3]); });
  var hasMan = {}; Object.keys(man).forEach(function (id) { var q = id.split('|'); hasMan[q[0] + '|' + q[1]] = 1; });
  wsl.forEach(function (w) {
    for (var off = 0; off <= 5; off++) {
      var t = new Date(base.getTime()); t.setDate(t.getDate() - off); var kk = kitKey_(t), id = w + '|' + kk;
      if (calc.has[id] || hasMan[id]) {
        KIT_CAPS.forEach(function (cp) { var mk = id + '|' + cp; out[w + '|' + cp] = man[mk] !== undefined ? man[mk] : slotBal_(calc.slots[w + '|' + cp], kk); });
        info[w] = { date: kitDs_(kk), src: calc.has[id] ? 'kitty' : 'admin' }; return;
      }
    }
    info[w] = null;
  });
  return { v: out, info: info, ly: kitDs_(kitKey_(base)) };
}
function kitView_(u, date, wsl) {
  var d = kitDate_(date), calc = kitCalc_(wsl), pm = kitPrevMap_(d.k, wsl, calc), rows = [], upd = {};
  wsl.forEach(function (w) {
    KIT_CAPS.forEach(function (cp) {
      var x = calc.slots[w + '|' + cp], dd = (x && x.day[d.k]) || { issue: 0, repair: 0, iso: 0, rfo: 0 };
      var o = { ws: w, cap: cp, open: slotOpen_(x, d.k), issue: dd.issue, repair: dd.repair, iso: dd.iso, rfo: dd.rfo, prev: pm.v[w + '|' + cp] || 0 };
      o.bal = o.open + o.repair - o.issue - o.iso + o.rfo; rows.push(o);
      if (dd.on && !upd[w]) upd[w] = { on: dd.on, by: dd.by };
    });
  });
  return { date: d.s, caps: KIT_CAPS, wsList: wsl, rows: rows, updated: upd, ly: pm.ly, prevInfo: pm.info };
}
function kitVisible_(u) { return ws_().filter(function (w) { return inScope_(u, w); }); }
function kitReqList_(u, wsl) {
  var ok = {}; wsl.forEach(function (w) { ok[w] = 1; });
  return kitRows_(KR, KR_HDR).filter(function (r) { return String(r[6]) === 'PENDING' && ok[wsKey_(r[2])]; }).map(function (r) {
    var v = {}; try { v = JSON.parse(r[5]); } catch (e) {}
    return { id: String(r[0]), on: String(r[1]), ws: wsKey_(r[2]), date: String(r[3]), by: String(r[4]), vals: v };
  });
}
function getKitty(token, date) {
  var u = auth_(token), vis = kitVisible_(u), out = kitView_(u, date, vis);
  out.edit = (u.admin || !u.canEdit) ? [] : vis; out.admin = !!u.admin;
  out.pending = {}; kitReqList_(u, vis).forEach(function (r) { out.pending[r.ws] = r; });
  return out;
}
function saveKitty(token, ws, date, items) {
  var u = auth_(token), w = wsKey_(ws), d = kitDate_(date);
  if (!w) throw new Error('Unknown workshop.');
  if (u.admin || !u.canEdit || !inScope_(u, w)) throw new Error('Only the workshop can fill its own kitty.');
  var map = {}; (items || []).forEach(function (x) { var cp = +x.cap; if (KIT_CAPS.indexOf(cp) >= 0) map[cp] = x; });
  var l = lock_();
  try {
    var mine = KIT_CAPS.map(function (cp) { var x = map[cp] || {}; return [cp, kitNum_(x.issue), kitNum_(x.repair), kitNum_(x.iso), kitNum_(x.rfo)]; });
    kdsSetDay_(d.k, w, mine, kitNow_(), u.name);                       // only the cell of this day is written
  } finally { l.releaseLock(); }
  return getKitty(token, d.s);
}
// Division Head types the balance kitty of last year (until the app holds a full year of its own data)
function getKittyPrev(token, date) {
  var u = auth_(token); if (!u.admin) throw new Error('Only the Division Head can fill last year kitty.');
  var d = kitDate_(date), vis = kitVisible_(u), vals = {}, ok = {}; vis.forEach(function (w) { vals[w] = {}; ok[w] = 1; });
  kitRows_(KP, KP_HDR).forEach(function (r) { var w = wsKey_(r[1]); if (ok[w] && dk_(r[0]) === d.k) vals[w][+r[2]] = kitNum_(r[3]); });
  return { date: d.s, caps: KIT_CAPS, wsList: vis, vals: vals };
}
function kitSavePrev(token, date, data) {
  var u = auth_(token); if (!u.admin) throw new Error('Only the Division Head can fill last year kitty.');
  var d = kitDate_(date), vis = kitVisible_(u), l = lock_();
  try {
    var keep = kitRows_(KP, KP_HDR).filter(function (r) { return !(dk_(r[0]) === d.k && vis.indexOf(wsKey_(r[1])) >= 0) && String(r[0]).trim() !== ''; }), now = kitNow_();
    vis.forEach(function (w) {
      var m = (data || {})[w]; if (!m || !KIT_CAPS.some(function (cp) { return kitNum_(m[cp]) > 0; })) return;      // an all-zero table means 'not filled'
      KIT_CAPS.forEach(function (cp) { keep.push([d.s, w, cp, kitNum_(m[cp]), u.name, now]); });
    });
    kitWrite_(KP, KP_HDR, keep);
  } finally { l.releaseLock(); }
  return getKittyPrev(token, d.s);
}
// the workshop asks to change the opening kitty of a date -> the Division Head approves
function kitRequestOpen(token, ws, date, vals) {
  var u = auth_(token), w = wsKey_(ws), d = kitDate_(date);
  if (!w) throw new Error('Unknown workshop.');
  if (u.admin || !u.canEdit || !inScope_(u, w)) throw new Error('Only the workshop can request a change of its opening kitty.');
  var clean = {}, any = false; KIT_CAPS.forEach(function (cp) { if (vals && vals[cp] !== undefined && vals[cp] !== '') { clean[cp] = kitNum_(vals[cp]); any = true; } });
  if (!any) throw new Error('Enter the new opening kitty.');
  var l = lock_();
  try {
    var rows = kitRows_(KR, KR_HDR).map(function (r) { if (String(r[6]) === 'PENDING' && wsKey_(r[2]) === w) { r[6] = 'REPLACED'; } return r; });
    rows.push([Utilities.getUuid().slice(0, 8), kitNow_(), w, d.s, u.name, JSON.stringify(clean), 'PENDING', '', '']);
    kitWrite_(KR, KR_HDR, rows);
  } finally { l.releaseLock(); }
  return getKitty(token, d.s);
}
function kitCancelReq(token, ws, date) {
  var u = auth_(token), w = wsKey_(ws);
  if (!w || u.admin || !inScope_(u, w)) throw new Error('Not allowed.');
  var l = lock_();
  try { kitWrite_(KR, KR_HDR, kitRows_(KR, KR_HDR).map(function (r) { if (String(r[6]) === 'PENDING' && wsKey_(r[2]) === w) { r[6] = 'CANCELLED'; r[8] = kitNow_(); } return r; })); }
  finally { l.releaseLock(); }
  return getKitty(token, date);
}
// requests waiting for the Division Head (shown as a small popup)
function kitPending(token) {
  var u = auth_(token); if (!u.admin) return [];
  var vis = kitVisible_(u), list = kitReqList_(u, vis);
  list.forEach(function (r) {
    var cur = {}, v = kitView_(u, r.date, [r.ws]); v.rows.forEach(function (x) { cur[x.cap] = x.open; }); r.cur = cur;
  });
  return list;
}
function kitDecide(token, id, approve) {
  var u = auth_(token); if (!u.admin) throw new Error('Only the Division Head can approve.');
  var l = lock_();
  try {
    var reqs = kitRows_(KR, KR_HDR), hit = null;
    reqs.forEach(function (r) { if (String(r[0]) === String(id) && String(r[6]) === 'PENDING') hit = r; });
    if (!hit) throw new Error('This request is no longer pending.');
    if (approve) {
      var w = wsKey_(hit[2]), vals = {}; try { vals = JSON.parse(hit[5]); } catch (e) {}
      var now = kitNow_(), keep = kitRows_(KO, KO_HDR).filter(function (r) { return !(String(r[0]) === String(hit[3]) && wsKey_(r[1]) === w && vals[+r[2]] !== undefined); });
      Object.keys(vals).forEach(function (cp) { keep.push([hit[3], w, +cp, kitNum_(vals[cp]), String(hit[4]), now]); });
      kitWrite_(KO, KO_HDR, keep);
    }
    hit[6] = approve ? 'APPROVED' : 'REJECTED'; hit[7] = u.name; hit[8] = kitNow_();
    kitWrite_(KR, KR_HDR, reqs);
  } finally { l.releaseLock(); }
  return kitPending(token);
}

// summary ke kisi number par click -> poori list
function oldList_(u, c) {                                 // read-only list of the old Google Sheets (PR date between From and To)
  var out = oldRead_(c.F).filter(function (a) { var d = dk_(a[I_PRD]); return d >= c.F && d <= c.T && inScope_(u, wsKey_(a[I_WS])); });
  out.sort(function (a, b) { return (dk_(a[I_PRD]) - dk_(b[I_PRD])) || (a[I_PR] < b[I_PR] ? -1 : 1); });
  return { total: out.length, rows: out.slice(0, 2000).map(function (a) { return a.concat([0]); }) };
}
function getList(token, spec) {
  var u = auth_(token), c = ctx_(spec), m = +spec.metric;
  if (spec.old) return oldList_(u, c);
  var dv = spec.divs ? spec.divs.reduce(function (o, k) { o[k] = 1; return o; }, {}) : null;
  var sbf = spec.subs ? spec.subs.reduce(function (o, k) { o[k] = 1; return o; }, {}) : null;
  var cp = spec.caps ? spec.caps.reduce(function (o, k) { o[k] = 1; return o; }, {}) : null, prd = spec.prd ? dk_(spec.prd) : 0;
  for (var t = 0; t < 2; t++) {
    var out = load_(u, c).filter(function (r) {
      if (spec.ws && r.ws !== spec.ws) return false;
      if (dv && !dv[r.div]) return false;
      if (sbf && !sbf[r.sub]) return false;
      if (cp && !cp[r.cap]) return false;
      if (prd && r.prd !== prd) return false;
      if (spec.ph && r.prd > +spec.ph) return false;
      if (spec.pl && r.prd < +spec.pl) return false;
      return flags_(r, c)[m];
    });
    out.sort(function (a, b) { return (a.prd - b.prd) || (a.pr < b.pr ? -1 : 1); });
    var top = out.slice(0, 2000);
    // poori row sirf unhi ki padhi jaati hai jo list me dikhengi (index me sirf 8 cheezein hain)
    var got = fetchRows_(idx_(), top.filter(function (r) { return !r.arch; }).map(function (r) { return r.i; }));
    if (!got) { bump_(); continue; }                       // sheet index se mel nahi khayi -> index naya banao, dobara
    return {
      total: out.length,
      rows: top.map(function (r) { var a = r.arch ? r.a : got[r.i]; return a.concat([(!r.arch && canEdit_(u, a)) ? 1 : 0]); })
    };
  }
  throw new Error('The data is changing right now, please try again.');
}
function getCols(token) { auth_(token); return COLS; }

/* ---------------- date range download (Data tab) ---------------- */
// basis: any = PR/Issue/Replacement me se koi bhi date From-To ke beech ho, pr / issue / rep = sirf wahi date
function getDataRange(token, p) {
  var u = auth_(token), F = dk_(p.from), T = dk_(p.to);
  if (!F || !T) throw new Error('Enter the From / To dates in DD.MM.YYYY format.');
  if (F > T) throw new Error('The "From" date cannot be later than the "To" date.');
  var basis = String(p.basis || 'any'), w = wsKey_(p.ws);
  var which = basis === 'pr' ? [0] : basis === 'issue' ? [1] : basis === 'rep' ? [2] : basis === 'ret' ? [3] : [0, 1, 2, 3];
  var out = [];
  var src = readAll_();
  if (F < arcLimit_()) src = src.concat(archRead_(F));            // purana data ARCHIVE se
  src.forEach(function (a) {
    var wk = wsKey_(a[I_WS]);
    if (!inScope_(u, wk) || (w && wk !== w)) return;
    var d = [dk_(a[I_PRD]), dk_(a[I_ISS]), dk_(a[I_REP]), dk_(a[I_RET])], hit = false;
    which.forEach(function (i) { if (d[i] >= F && d[i] <= T) hit = true; });
    if (hit) out.push({ a: a, k: d[0] });
  });
  out.sort(function (x, y) { return (x.k - y.k) || (x.a[I_PR] < y.a[I_PR] ? -1 : 1); });
  return { total: out.length, rows: out.map(function (o) { return o.a; }) };
}

/* ---------------- Driver Master (tab: Drivers) ---------------- */
// Columns: ID | Firm Name | Workshop | Driver Name | Mobile No. | Vehicle GPS URL | Vehicle No.
function drvSheet_() {
  var ss = ss_(), sh = ss.getSheetByName(DRV);
  if (!sh) {
    sh = ss.insertSheet(DRV);
    sh.getRange(1, 1, 1, 7).setValues([['ID', 'Firm Name', 'Workshop', 'Driver Name', 'Mobile No.', 'Vehicle GPS URL', 'Vehicle No.']]).setFontWeight('bold');
    sh.setFrozenRows(1);
    if (sh.getMaxColumns() > 7) sh.deleteColumns(8, sh.getMaxColumns() - 7);
    if (sh.getMaxRows() > 50) sh.deleteRows(51, sh.getMaxRows() - 50);
  }
  return sh;
}
function dn_(v) { return String(v == null ? '' : v).replace(/\u00a0/g, ' ').replace(/[\r\n\t]+/g, ' ').trim(); }
function nk_(s) { return String(s == null ? '' : s).toUpperCase().replace(/[^A-Z0-9]/g, ''); }
function mob_(s) {
  var d = String(s || '').replace(/\D/g, '');
  if (d.length === 12 && d.indexOf('91') === 0) d = d.slice(2); else if (d.length === 11 && d.charAt(0) === '0') d = d.slice(1);
  return d;
}
function drvKey_(r) { return wsKey_(r[2]) + '|' + nk_(r[1]) + '|' + nk_(r[3]); }
function drvAll_() {
  var sh = drvSheet_(), n = sh.getLastRow(), rows = [], fixed = false;
  if (n >= 2) rows = sh.getRange(2, 1, n - 1, 7).getValues().map(function (r) { return r.map(dn_); })
    .filter(function (r) { return r.slice(1).some(function (x) { return x !== ''; }); });
  rows.forEach(function (r) { if (!r[0]) { r[0] = 'd' + rnd_(8); fixed = true; } });
  rows.fixed = fixed;
  return rows;
}
function drvWrite_(rows) {
  var sh = drvSheet_(), n = sh.getLastRow();
  if (n >= 2) sh.getRange(2, 1, n - 1, 7).clearContent();
  if (!rows.length) return;
  var need = rows.length + 1;
  if (need > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), need - sh.getMaxRows());
  sh.getRange(2, 1, rows.length, 7).setNumberFormat('@').setValues(rows);
}
function dupRows_(rows) {                                 // every workshop gets the default firm DUPLICATE/WRONG PR (returns true when rows were added)
  var have = {}, add = false;
  rows.forEach(function (r) { if (isDup_(r[1])) { have[wsKey_(r[2])] = 1; if (r[3] === 'N/A' || r[3] !== DUP_DRV || r[4]) { r[3] = DUP_DRV; r[4] = ''; add = true; } } });   // default entry: firm DUPLICATE/WRONG PR, driver JE(D); name and mobile are typed by the user in each record
  ws_().forEach(function (w) { if (!have[w]) { rows.push(['d' + rnd_(8), DUP_FIRM, w, DUP_DRV, '', '', '']); add = true; } });
  return add;
}
function getDrivers(token) {
  var u = auth_(token), rows = drvAll_();
  if (u.canEdit && dupRows_(rows)) rows.fixed = true;
  if (rows.fixed && u.canEdit) { var l = lock_(); try { drvWrite_(rows); } finally { l.releaseLock(); } }   // purani manual entries ko ID de do
  var out = rows.filter(function (r) { return inScope_(u, wsKey_(r[2])); });
  return { rows: out, canEdit: u.canEdit, one: oneWs_(u), total: out.length };
}
// items = [{id: '' (nayi row) ya purani ID, r: [firm, workshop, driver, mobile, gps, vehicle]}]
function saveDrivers(token, items) {
  var u = auth_(token);
  if (!u.canEdit) throw new Error('You do not have permission to change the Driver Master.');
  if (!items || !items.length) return { added: 0, updated: 0, skipped: 0, notes: [] };
  var l = lock_();
  try {
    var rows = drvAll_(), pos = {}, kpos = {}, one = oneWs_(u), added = 0, updated = 0, skipN = 0, notes = [], bad = [], fin = {};
    rows.forEach(function (r, i) { pos[r[0]] = i; kpos[drvKey_(r)] = i; });
    function skip(n, who, why) { skipN++; bad.push(n - 1); if (notes.length < 15) notes.push('Line ' + n + ' (' + (who || '-') + '): ' + why); }
    items.forEach(function (it, n) {
      var s = it.r || [], firm = dn_(s[0]), ws = dn_(s[1]), drv = dn_(s[2]), mob = mob_(s[3]), gps = dn_(s[4]), veh = dn_(s[5]);
      if (!ws && one) ws = one;
      var who = drv || firm, why = '';
      if (!firm) why = 'Firm Name is empty';
      else if (!drv) why = 'Driver Name is empty';
      else if (!wsKey_(ws)) why = 'Workshop is empty/not recognised (AZAMGARH / MAU / BALLIA)';
      else if (!inScope_(u, wsKey_(ws))) why = 'This is not an entry of your workshop';
      else if (mob && mob.length !== 10) why = 'Mobile No. must be 10 digits';
      if (!why && isDup_(firm)) why = 'The firm DUPLICATE/WRONG PR is a default entry and cannot be added or changed';
      if (!why && it.id && pos[it.id] !== undefined && isDup_(rows[pos[it.id]][1])) why = 'The default entry DUPLICATE/WRONG PR cannot be changed';
      if (why) return skip(n + 1, who, why);
      var nr = ['', firm, wsKey_(ws), drv, mob, gps, veh], k = drvKey_(nr), old = it.id ? pos[it.id] : undefined;
      if (it.id && old === undefined) return skip(n + 1, who, 'the old entry was not found (did someone delete it?)');
      if (old !== undefined) {
        if (!inScope_(u, wsKey_(rows[old][2]))) return skip(n + 1, who, 'You do not have permission to change this entry');
        if (kpos[k] !== undefined && kpos[k] !== old) return skip(n + 1, who, 'This Firm + Driver already exists');
        nr[0] = rows[old][0]; delete kpos[drvKey_(rows[old])]; rows[old] = nr; kpos[k] = old; updated++; fin[n] = k;
      } else if (kpos[k] !== undefined) {                                      // pehle se hai -> update
        var ex = rows[kpos[k]], ch = false;
        [4, 5, 6].forEach(function (j) { if (nr[j] !== '' && nr[j] !== ex[j]) { ex[j] = nr[j]; ch = true; } });
        if (ch) updated++;
        fin[n] = k;
      } else {
        nr[0] = 'd' + rnd_(8); pos[nr[0]] = rows.length; kpos[k] = rows.length; rows.push(nr); added++; fin[n] = k;
      }
    });
    if (added || updated) drvWrite_(rows);
    var out = {}; Object.keys(fin).forEach(function (n) { var p = kpos[fin[n]]; if (p !== undefined) out[n] = rows[p]; });
    return { added: added, updated: updated, skipped: skipN, notes: notes, bad: bad, rows: out };
  } finally { l.releaseLock(); }
}
function deleteDrivers(token, ids) {
  var u = auth_(token);
  if (!u.canEdit) throw new Error('You do not have permission to delete drivers.');
  var l = lock_();
  try {
    var rows = drvAll_(), gone = {};
    (ids || []).forEach(function (i) { gone[i] = 1; });
    var keep = rows.filter(function (r) {
      if (!gone[r[0]]) return true;
      if (isDup_(r[1])) return true;                                   // the default DUPLICATE/WRONG PR entry stays
      if (!inScope_(u, wsKey_(r[2]))) throw new Error('You do not have permission to delete this entry.');
      return false;
    });
    drvWrite_(keep);
    return true;
  } finally { l.releaseLock(); }
}

/* ---------------- sheet chhota karna (bachi hui khali rows / columns hata deta hai) ---------------- */
// Google Sheet me 1000 rows x 26 columns pehle se hote hain; use hatane se file halki rehti hai aur
// 1 Cr cell ki limit door rehti hai. Naya data aane par rows apne aap sirf zarurat jitni judti hain.
function compactSheets(silent) {
  var ss = ss_(), out = [];
  [DATA, USERS, DRV, ARC].forEach(function (nm) {
    var sh = ss.getSheetByName(nm); if (!sh) return;
    var before = sh.getMaxRows() * sh.getMaxColumns();
    if (nm !== USERS) {                                               // beech ki poori khali rows hatao
      var lr0 = sh.getLastRow(), lc0 = sh.getLastColumn();
      if (lr0 > 2 && lc0 > 0) {
        var vals = sh.getRange(2, 1, lr0 - 1, lc0).getValues(), del = [];
        vals.forEach(function (r, i) { if (r.every(function (v) { return v === '' || v === null; })) del.push(i + 2); });
        for (var i = del.length - 1; i >= 0;) {
          var j = i; while (j > 0 && del[j - 1] === del[j] - 1) j--;
          sh.deleteRows(del[j], del[i] - del[j] + 1); i = j - 1;
        }
      }
    }
    var lr = Math.max(sh.getLastRow(), 2), lc = Math.max(sh.getLastColumn(), 1);
    if (sh.getMaxRows() > lr) sh.deleteRows(lr + 1, sh.getMaxRows() - lr);
    if (sh.getMaxColumns() > lc) sh.deleteColumns(lc + 1, sh.getMaxColumns() - lc);
    out.push(nm + ': ' + before + ' cells -> ' + sh.getMaxRows() * sh.getMaxColumns() + ' cells');
  });
  bump_();                                   // rows hat sakti hain -> tez index naya banana hoga
  if (!silent) alert_('Sheet chhoti ho gayi.\n' + out.join('\n'));
  return out;
}

/* ---------------- ARCHIVE: 2 saal se purana poora data, 1 record = 1 line = 1 cell ---------------- */
// DATA me ek record 36 cells leta hai. ARCHIVE me wahi record sirf 3 cells leta hai:
//   A = Key (WORKSHOP|PR NO)   B = record ki sabse nayi date (YYYYMMDD)   C = poora record (packed text)
// ARCHIVE hamesha B ke hisaab se sorted rehta hai, isliye purani date ki summary me sirf zaruri hissa hi padhta hai.
// FINANCIAL YEAR (India: 1 April - 31 March): DATA tab me sirf CURRENT financial year ka data rehta hai (+ har wo record jo abhi poora nahi hua).
// Pichhle financial years ke poore ho chuke records ARCHIVE me jate hain; Comparison / purani date ki summary me ARCHIVE apne aap padha jata hai.
// Sirf wahi record archive hote hain jo poori tarah band ho chuke hon (wrong PR, ya RETURN DATE bhari ho) aur jinki koi bhi date current FY ki 1 April se pehle ki ho.
var ARCH_FY = true;                                     // true = financial year (April-March);  false = purana tarika (ARCH_YEARS saal)
function fyStart_() {                                   // current financial year ki 1 April, YYYYMMDD
  var n = new Date(), y = +Utilities.formatDate(n, TZ, 'yyyy'), m = +Utilities.formatDate(n, TZ, 'MM');
  return (m >= 4 ? y : y - 1) * 10000 + 401;
}
function cutoff2y_() {
  var n = new Date();
  return (+Utilities.formatDate(n, TZ, 'yyyy') - ARCH_YEARS) * 10000 + (+Utilities.formatDate(n, TZ, 'MM')) * 100 + (+Utilities.formatDate(n, TZ, 'dd'));
}
// MONTH mode (after the Admin pressed "Convert"): DATA keeps only the current month + the previous month (+ every unfinished record)
function monthCut_() {
  var n = new Date(), y = +Utilities.formatDate(n, TZ, 'yyyy'), m = +Utilities.formatDate(n, TZ, 'MM') - 1;
  if (m < 1) { m = 12; y--; }
  return y * 10000 + m * 100 + 1;
}
function ymNow_() { return Utilities.formatDate(new Date(), TZ, 'yyyyMM'); }
function conv_() { try { return !!cfg_().conv; } catch (e) { return false; } }
function cutoff_() { return conv_() ? monthCut_() : (ARCH_FY ? fyStart_() : cutoff2y_()); }
// kya ye record DATA se ARCHIVE me ja sakta hai?
function arcOk_(r, cut) {
  var mx = maxDate_(r);
  if (!(mx > 0 && mx < cut) || !dk_(r[I_PRD])) return false;
  if (isDup_(r[I_FIRM]) || dk_(r[I_RET])) return true;                       // wrong PR ya return ho chuka = poora band
  return !!(dk_(r[I_ISS]) && dk_(r[I_REP])) && mx < cutoff2y_();              // replace ho gaya par RETURN DATE nahi bhari: "Pending for return" me gine jate hain, isliye 2 saal baad hi ARCHIVE
}
// ek sheet ke DATA rows + (agar minK purani FY me ho to) ARCHIVE ke rows - DISCOM head ke Comparison / Summary ke liye
function rowsWithArc_(sid, minK) {
  var rows = sheetRows_(sid);
  try { useSheet_(sid); if (minK && minK < arcLimit_()) rows = rows.concat(archRead_(minK)); } catch (e) {}
  return rows;
}
function maxDate_(r) {                       // record ki sabse nayi date; koi date galat ho to -1 (archive nahi karenge)
  var mx = 0;
  for (var i = 0; i < DATE_IDX.length; i++) {
    var v = r[DATE_IDX[i]]; if (v === '') continue;
    var d = dk_(v); if (!d) return -1; if (d > mx) mx = d;
  }
  return mx;
}
function archSheet_() {
  var ss = ss_(), sh = ss.getSheetByName(ARC);
  if (!sh) {
    sh = ss.insertSheet(ARC);
    sh.getRange(1, 1, 1, 3).setValues([['Key (Workshop|PR NO)', 'Last date (YYYYMMDD)', 'Record (packed)']]).setFontWeight('bold');
    sh.setFrozenRows(1);
    if (sh.getMaxColumns() > 3) sh.deleteColumns(4, sh.getMaxColumns() - 3);
    if (sh.getMaxRows() > 50) sh.deleteRows(51, sh.getMaxRows() - 50);
  }
  return sh;
}
var ARCH_IX = null;
// ---- old Google Sheets (after the data was moved to a new sheet) ----
function oldCut_() { var m = 0; (cfg_().old || []).forEach(function (o) { if (+o.cut > m) m = +o.cut; }); return m; }
function arcLimit_() { return Math.max(cutoff_(), oldCut_()); }
function migCut_() { var d = new Date(Date.now() - 365 * 86400000); return +Utilities.formatDate(d, TZ, 'yyyy') * 10000 + (+Utilities.formatDate(d, TZ, 'MM')) * 100 + (+Utilities.formatDate(d, TZ, 'dd')); }
function oldElig_(r, cut) {                           // a finished record whose last date is older than the cut: it lives only in the old sheet
  if (r[I_PR] === '') return false; var mx = maxDate_(r);
  return mx > 0 && mx < cut && !!dk_(r[I_PRD]) && ((dk_(r[I_ISS]) && dk_(r[I_REP])) || !!dk_(r[I_RET]) || isDup_(r[I_FIRM])) && !!wsKey_(r[I_WS]);
}
var OLD_MEM = null;
function oldRead_(minK) {
  if (!OLD_MEM || OLD_MEM.sid !== SS_ID) OLD_MEM = { sid: SS_ID, k: {} };
  var out = [];
  (cfg_().old || []).forEach(function (o) {
    var all = OLD_MEM.k[o.sid];
    if (!all) {
      all = []; var src; try { src = SpreadsheetApp.openById(o.sid); } catch (e) { OLD_MEM.k[o.sid] = all; return; }
      var d = src.getSheetByName(DATA), n = d ? d.getLastRow() : 0, cut = +o.cut;
      if (n >= 2) d.getRange(2, 1, n - 1, NC).getValues().forEach(function (r) { r = r.map(function (v, i) { return norm_(i, v); }); if (oldElig_(r, cut)) all.push(r); });
      var a = src.getSheetByName(ARC), m = a ? a.getLastRow() : 0;
      if (m >= 2) a.getRange(2, 3, m - 1, 1).getValues().forEach(function (r) { try { all.push(unpack_(r[0])); } catch (e) {} });
      OLD_MEM.k[o.sid] = all;
    }
    all.forEach(function (r) { if (maxDate_(r) >= minK) out.push(r); });
  });
  return out;
}
function archIdx_() {                        // key -> sheet row number
  if (ARCH_IX) return ARCH_IX;
  ARCH_IX = {};
  var sh = ss_().getSheetByName(ARC), n = sh ? sh.getLastRow() : 0;
  if (n >= 2) sh.getRange(2, 1, n - 1, 1).getValues().forEach(function (r, i) { if (r[0] !== '') ARCH_IX[String(r[0])] = i + 2; });
  return ARCH_IX;
}
function unpack_(txt) {
  var a = JSON.parse(txt), r = [];
  for (var i = 0; i < NC; i++) r.push(norm_(i, a[i]));
  return r;
}
// minK (YYYYMMDD) ya usse nayi last-date wale archived records
function archRead_(minK) {
  var out = [], sh = ss_().getSheetByName(ARC), n = sh ? sh.getLastRow() : 0;
  if (n >= 2) {
    var b = sh.getRange(2, 2, n - 1, 1).getValues(), st = -1;
    for (var i = 0; i < b.length; i++) if (+b[i][0] >= minK) { st = i; break; }
    if (st >= 0) sh.getRange(st + 2, 3, n - 1 - st, 1).getValues().forEach(function (r) { try { out.push(unpack_(r[0])); } catch (e) {} });
  }
  return out.concat(oldRead_(minK));                      // plus the old Google Sheets (if the data was moved to a new sheet)
}
// kisi purani PR par naya data paste/save ho to record ARCHIVE se nikal kar DATA me wapas aa jata hai
function restore_(key, all, pos, added, rs) {
  var rn = archIdx_()[key]; if (!rn) return;
  var rec; try { rec = unpack_(archSheet_().getRange(rn, 3).getValue()); } catch (e) { return; }
  pos[key] = all.length; all.push(rec); added.push(rec); rs.push(rn); delete ARCH_IX[key];
  try { statsAdd_([rec], -1); } catch (e) {}
}
function archDel_(rowNums) {
  var sh = archSheet_(), nums = rowNums.slice().sort(function (a, b) { return b - a; });
  nums.forEach(function (n) { sh.deleteRow(n); });
  ARCH_IX = null;
}
/* ---------------- STATS: har mahine ka chhota hisaab (sirf ARCHIVE ke records ka) ----------------
 * Comparison ko pichhle saal ke records ek-ek karke nahi padhne padte. STATS tab me ARCHIVE ke records ka
 * "PR date x workshop x division x substation x capacity -> kitne PR" ka count rehta hai (non-duplicate PR).
 * Row = Month (YYYYMM) | Workshop | Part | counts (ek cell). Sirf "Convert" ke baad bana hota hai; usse pehle purana tarika chalta hai. */
var ST = 'STATS', ST_HDR = ['Month (YYYYMM)', 'Workshop', 'Part', 'Counts (1 workshop-month = 1 cell)'];
var STATS_MEM = null;
function statsTab_() { return ss_().getSheetByName(ST); }
function statsReady_() { try { return conv_() && !(cfg_().old || []).length && !!statsTab_(); } catch (e) { return false; } }
function statsLoad_() {
  if (STATS_MEM && STATS_MEM.sid === SS_ID) return STATS_MEM.m;
  var m = {}, sh = statsTab_(), n = sh ? sh.getLastRow() : 0;
  if (n >= 2) sh.getRange(2, 1, n - 1, 4).getValues().forEach(function (r) {
    var ym = String(r[0]).trim(), w = String(r[1]).trim(); if (!ym || !w) return;
    var o; try { o = JSON.parse(r[3]); } catch (e) { return; }
    var t = (m[ym] || (m[ym] = {}))[w] || (m[ym][w] = {});
    for (var k in o) t[k] = (t[k] || 0) + o[k];
  });
  STATS_MEM = { sid: SS_ID, m: m };
  return m;
}
function statsSave_(m) {
  var sh = statsTab_() || ss_().insertSheet(ST);
  if (sh.getLastRow() < 1) { sh.getRange(1, 1, 1, 4).setValues([ST_HDR]).setFontWeight('bold'); sh.setFrozenRows(1); }
  var rows = [];
  Object.keys(m).sort().forEach(function (ym) {
    Object.keys(m[ym]).sort().forEach(function (w) {
      var o = m[ym][w], keys = Object.keys(o), part = {}, len = 0, pn = 1;
      keys.forEach(function (k) {
        if (len + k.length + 12 > 40000) { rows.push([ym, w, pn++, JSON.stringify(part)]); part = {}; len = 0; }
        part[k] = o[k]; len += k.length + 12;
      });
      if (Object.keys(part).length) rows.push([ym, w, pn, JSON.stringify(part)]);
    });
  });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 4).clearContent();
  if (rows.length) {
    if (sh.getMaxRows() < rows.length + 1) sh.insertRowsAfter(sh.getMaxRows(), rows.length + 1 - sh.getMaxRows());
    sh.getRange(2, 1, rows.length, 4).setNumberFormat('@').setValues(rows);
  }
  STATS_MEM = { sid: SS_ID, m: m };
}
// sign = +1 (records went to ARCHIVE) / -1 (a record came back to DATA)
function statsApply_(m, recs, sign) {
  recs.forEach(function (r) {
    if (isDup_(r[I_FIRM])) return;
    var prd = dk_(r[I_PRD]), w = wsKey_(r[I_WS]); if (!prd || !w) return;
    var ym = String(Math.floor(prd / 100)), t = (m[ym] || (m[ym] = {}))[w] || (m[ym][w] = {});
    var k = [prd, divKey_(r[I_DIV]), divKey_(r[I_SUB]), capOf_(r[I_CAP])].join('\t');
    t[k] = (t[k] || 0) + sign; if (t[k] <= 0) delete t[k];
  });
}
function statsAdd_(recs, sign) {
  if (!conv_() || !statsTab_() || !recs.length) return;
  var m = statsLoad_(); statsApply_(m, recs, sign); statsSave_(m);
}
function statsRebuild_() {
  var sh = archSheet_(), n = sh.getLastRow(), m = {}, recs = [];
  if (n >= 2) sh.getRange(2, 3, n - 1, 1).getValues().forEach(function (r) { try { recs.push(unpack_(r[0])); } catch (e) {} });
  statsApply_(m, recs, 1); statsSave_(m);
  return recs.length;
}
// fn(ws, div, sub, cap, prd, count) for every archived PR whose PR date is between lo and hi
function statsIter_(lo, hi, fn) {
  var m = statsLoad_();
  Object.keys(m).forEach(function (ym) {
    if (+ym * 100 + 31 < lo || +ym * 100 > hi) return;
    Object.keys(m[ym]).forEach(function (w) {
      var t = m[ym][w];
      for (var k in t) { var q = k.split('\t'), d = +q[0]; if (d >= lo && d <= hi) fn(w, q[1], q[2], +q[3], d, t[k]); }
    });
  });
}

// cutOv = cutoff to use (Convert uses the month cutoff before the flag is set)
function archiveCore_(silent, cutOv) {
  var sh = ss_().getSheetByName(DATA); if (!sh) { if (!silent) alert_('The DATA tab does not exist -> run Setup.'); return 0; }
  var cut = cutOv || cutoff_(), all = readAll_(true), keep = [], move = [], ix = archIdx_();
  all.forEach(function (r) {
    if (r[I_PR] === '') { keep.push(r); return; }
    if (arcOk_(r, cut) && wsKey_(r[I_WS]) && !ix[rk_(r[I_WS], r[I_PR])]) move.push(r);
    else keep.push(r);
  });
  if (!move.length) { if (!silent) alert_('No old data eligible for archiving was found.\n(finished records of earlier months only)'); return 0; }
  var ash = archSheet_(), start = Math.max(ash.getLastRow(), 1) + 1, seen = {};
  var out = []; move.forEach(function (r) { var k = rk_(r[I_WS], r[I_PR]); if (seen[k]) return; seen[k] = 1; out.push([k, String(maxDate_(r)), JSON.stringify(r)]); });
  var need = start + out.length - 1;
  if (need > ash.getMaxRows()) ash.insertRowsAfter(ash.getMaxRows(), need - ash.getMaxRows());
  for (var i = 0; i < out.length; i += 4000) { var part = out.slice(i, i + 4000); ash.getRange(start + i, 1, part.length, 3).setNumberFormat('@').setValues(part); }
  SpreadsheetApp.flush();
  if (ash.getLastRow() < need) throw new Error('Writing the archive did not complete - nothing in DATA was changed.');
  ash.getRange(2, 1, ash.getLastRow() - 1, 3).sort({ column: 2, ascending: true });
  ARCH_IX = null;
  if (conv_()) statsAdd_(move, 1);
  // pehle bache hue records upar likho, phir neeche ka bacha hissa saaf karo (beech me fail ho to bhi data khona nahi)
  writeRows_(sh, 2, keep);
  if (all.length > keep.length) sh.getRange(keep.length + 2, 1, all.length - keep.length, NC).clearContent();
  bump_();
  compactSheets(true);
  if (!silent) alert_(out.length + ' old records were moved to ARCHIVE (cells: ' + out.length * NC + ' -> ' + out.length * 3 + ').\nThey will be included automatically for old dates in Summary / Excel download.');
  return out.length;
}
function archiveOld(silent) {
  silent = silent === true;
  var l = lock_();
  try { return archiveCore_(silent); } finally { l.releaseLock(); }
}
// A new month started: finished records of older months move from DATA to ARCHIVE (once per month, at the first login)
function monthRoll_() {
  var c = cfg_(); if (!c.conv || c.ym === ymNow_()) return;
  var l = lock_();
  try {
    if (cfgRead_().ym === ymNow_()) return;
    cfgSetQ_('ArchYM', ymNow_());
    archiveCore_(true);
    cfgSetQ_('LastRoll', bkNow_());
  } finally { l.releaseLock(); }
}
// Admin button "Convert": DATA keeps only current + previous month (+ unfinished records), the rest goes to ARCHIVE (1 record = 1 cell), STATS is built
function adminConvert(token) {
  var a = auth_(token); if (!a.admin) throw new Error('Admin only.');
  var l = lock_();
  try {
    var ss = ss_(), src = ss.getSheetByName(DATA); if (!src) throw new Error('The DATA tab does not exist.');
    var nm = 'DATA_backup_' + Utilities.formatDate(new Date(), TZ, 'yyyyMMdd'), bk = ss.getSheetByName(nm);
    if (!bk) { bk = src.copyTo(ss); bk.setName(nm); try { bk.hideSheet(); } catch (e) {} }
    var moved = archiveCore_(true, monthCut_());
    var n = statsRebuild_();
    cfgSetQ_('Converted', '1'); cfgSetQ_('ArchYM', ymNow_()); cfgSetQ_('LastConvert', bkNow_() + ' by ' + a.name);
    bump_();
    var d = ss.getSheetByName(DATA), hot = 0; if (d && d.getLastRow() > 1) hot = readAll_(false).length;
    return { moved: moved, hot: hot, arch: n, backup: nm };
  } finally { l.releaseLock(); }
}

/* =================================================================
 *  PEHLI BAAR SETUP  (Admin)  +  sheet connect check
 * ================================================================= */
// 'ready' = setup ho chuka, 'empty' = naya khali sheet, 'legacy' = pehle se kuch data/users hain
function state_(ss) {
  var cf = ss.getSheetByName(CONF);
  if (cf && cf.getLastRow() >= 1) {
    var v = cf.getRange(1, 1, cf.getLastRow(), 2).getValues();
    for (var i = 0; i < v.length; i++) if (String(v[i][0]).trim() === 'SetupDone' && String(v[i][1]).trim()) return 'ready';
  }
  var d = ss.getSheetByName(DATA), us = ss.getSheetByName(USERS);
  if ((d && d.getLastRow() > 1) || (us && us.getLastRow() > 1)) return 'legacy';
  return 'empty';
}
function checkSheet(link) {
  var id = parseId_(link);
  if (!id) return { ok: false, msg: 'This does not look like a Google Sheet link. Open the sheet and paste the full link from the address bar.' };
  var ss;
  try { ss = SpreadsheetApp.openById(id); } catch (e) { return { ok: false, msg: NOSHEET }; }
  useSheet_(id);
  var st = state_(ss), out = { ok: true, sid: id, state: st, title: ss.getName() };
  if (st === 'ready') { var c = cfg_(); out.org = c.org; out.workshops = c.workshops; }
  return out;
}
function mkTab_(ss, name, hdr, keep) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clear();
  sh.getRange(1, 1, 1, hdr.length).setValues([hdr]).setFontWeight('bold').setBackground('#e3e9fb');
  sh.setFrozenRows(1);
  if (sh.getMaxColumns() > hdr.length) sh.deleteColumns(hdr.length + 1, sh.getMaxColumns() - hdr.length);
  if (sh.getMaxRows() > 50) sh.deleteRows(51, sh.getMaxRows() - 50);
  return sh;
}
function build_(ss, org, ws, aid, adisp, apw, discom) {
  var dt = mkTab_(ss, DATA, COLS);
  dt.getRange(1, 1, dt.getMaxRows(), NC).setNumberFormat('@');           // dates text (DD.MM.YYYY) hi rahein
  dt.setColumnWidths(1, NC, 120);
  var us = mkTab_(ss, USERS, ['User ID', 'Password', 'Role', 'Display Name', 'Can Edit', 'Active', 'Scope', 'Admin', 'Must change']);
  us.getRange(2, 1, 1, 9).setNumberFormat('@');
  us.getRange(2, 1, 1, 9).setValues([[aid, mk_(apw), 'ADMIN', adisp || aid, true, true, 'ALL', true, false]]);
  var dr = mkTab_(ss, DRV, ['ID', 'Firm Name', 'Workshop', 'Driver Name', 'Mobile No.', 'Vehicle GPS URL', 'Vehicle No.']);
  dr.getRange(2, 1, dr.getMaxRows() - 1, 7).setNumberFormat('@');
  dr.getRange(2, 1, ws.length, 7).setValues(ws.map(function (w) { return ['d' + rnd_(8), DUP_FIRM, w, DUP_DRV, '', '', '']; }));      // default firm for wrong / duplicate PRs
  mkTab_(ss, ARC, ['Key (Workshop|PR NO)', 'Last date (YYYYMMDD)', 'Record (packed)']);
  var cf = mkTab_(ss, CONF, ['Setting', 'Value']);
  cf.getRange(2, 1, 7, 2).setNumberFormat('@');
  cf.getRange(2, 1, 7, 2).setValues([
    ['App', 'TF-Tracker'], ['Org', org || ''], ['Discom', discom || ''], ['Workshops', ws.join(',')], ['Version', '2'],
    ['Note', 'Change the workshop list from the Admin tab of the app. Always add a new workshop at the end.'],
    ['SetupDone', Utilities.formatDate(new Date(), TZ, 'dd.MM.yyyy HH:mm')]]);
  cf.setColumnWidths(1, 2, 220);
  us.setColumnWidths(1, 9, 110); dr.setColumnWidths(1, 7, 130);
  ss.getSheets().forEach(function (s) {                                    // default khali "Sheet1" hata do
    if ([DATA, USERS, DRV, ARC, CONF].indexOf(s.getName()) >= 0) return;
    if (/^(sheet\s*\d*)$/i.test(s.getName()) && s.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  try { ss.setActiveSheet(dt); ss.moveActiveSheet(1); } catch (e) {}
}

/* ---------------- Central registry (developer's Drive, created automatically) ----------------
 * Spreadsheet "TF-Tracker - Registry" with two tabs:
 *   Admins : one row per setup  (time, Google email, Profit Center = Admin ID, DISCOM, org, workshops, sheet title, sheet link, sheet id)
 *   Users  : User ID -> Sheet ID (so anybody can log in with just ID + password; also keeps every ID unique)
 * The Google email is verified by Google (the one-time access token is checked), it is not typed by the user.
 * To find it: run registryUrl() once in Apps Script and open the link in the Execution log. */
var ADM_HDR = ['Time', 'Admin Google email', 'Profit Center (Admin ID)', 'DISCOM', 'Organisation', 'Workshops', 'Sheet title', 'Sheet link', 'Sheet ID', 'Division (EWD)'];
function regSS_() {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty('REGISTRY_ID'), ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create('TF-Tracker - Registry');
    props.setProperty('REGISTRY_ID', ss.getId());
    var a = ss.getSheets()[0]; a.setName('Admins');
    a.getRange(1, 1, 1, ADM_HDR.length).setValues([ADM_HDR]).setFontWeight('bold'); a.setFrozenRows(1);
    var u = ss.insertSheet('Users');
    u.getRange(1, 1, 1, 4).setValues([['User ID', 'Sheet ID', 'Role', 'Created']]).setFontWeight('bold'); u.setFrozenRows(1);
  }
  return ss;
}
function registryUrl() { var u = regSS_().getUrl(); Logger.log(u); return u; }
function regTab_(n) { var ss = regSS_(); return ss.getSheetByName(n) || ss.insertSheet(n); }
function regUsers_() {                                   // { userid(lowercase): sheetId }
  var c = CacheService.getScriptCache(), s = c.get('reg_users');
  if (s) { try { return JSON.parse(s); } catch (e) {} }
  var sh = regTab_('Users'), n = sh.getLastRow(), m = {};
  if (n >= 2) sh.getRange(2, 1, n - 1, 2).getValues().forEach(function (r) { if (r[0]) if (!m[String(r[0]).trim().toLowerCase()]) m[String(r[0]).trim().toLowerCase()] = String(r[1]).trim(); });
  try { c.put('reg_users', JSON.stringify(m), 120); } catch (e) {}
  return m;
}
function regUserSids_(id) {                              // one ID can exist in several organisations (a JE/AE who was transferred)
  var k = String(id || '').trim().toLowerCase(), cc = CacheService.getScriptCache(), map = null, s = cc.get('reg_us2');      // ID -> sheets map is cached 10 min: login no longer opens the Registry sheet every time
  if (s) { try { map = JSON.parse(s); } catch (e) { map = null; } }
  if (!map) {
    map = {}; var sh = regTab_('Users'), n = sh.getLastRow();
    if (n >= 2) sh.getRange(2, 1, n - 1, 2).getValues().forEach(function (r) { var q = String(r[1]).trim(), u = String(r[0]).trim().toLowerCase(); if (!u || !q) return; var a = map[u] || (map[u] = []); if (a.indexOf(q) < 0) a.push(q); });
    try { cc.put('reg_us2', JSON.stringify(map), 600); } catch (e) {}
  }
  return (map[k] || []).slice();
}
function regAddUser_(id, sid, role) {
  if (regUserSids_(id).indexOf(String(sid).trim()) >= 0) return;
  regTab_('Users').appendRow([String(id).trim(), sid, role || '', Utilities.formatDate(new Date(), TZ, 'dd.MM.yyyy HH:mm')]);
  try { CacheService.getScriptCache().removeAll(['reg_users', 'reg_us2']); } catch (e) {}
}
function regAdmins_() {                                  // rows of the Admins tab
  var sh = regTab_('Admins'), n = sh.getLastRow();
  return n >= 2 ? sh.getRange(2, 1, n - 1, ADM_HDR.length).getValues() : [];
}
var GERR_ = '';
function googleEmail_(token) {
  GERR_ = '';
  if (!token) { GERR_ = 'no token'; return ''; }
  try {
    var r = UrlFetchApp.fetch('https://www.googleapis.com/oauth2/v3/userinfo', { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) { GERR_ = 'Google answered ' + r.getResponseCode(); return ''; }
    return String(JSON.parse(r.getContentText()).email || '');
  } catch (e) { GERR_ = String((e && e.message) || e).slice(0, 160); return ''; }
}
// Run this ONCE in the Apps Script editor (select authorizeAll -> Run -> Allow): it asks for every permission the app needs (e-mail check, e-mail OTP, daily backup trigger)
function authorizeAll() {
  UrlFetchApp.fetch('https://www.googleapis.com/oauth2/v3/tokeninfo?access_token=x', { muteHttpExceptions: true });
  MailApp.getRemainingDailyQuota();
  ScriptApp.getProjectTriggers();
  SpreadsheetApp.create('TF-Tracker permission check').getId();
  Logger.log('All permissions granted. Creator e-mail = ' + creatorEmail_() + '. Now Deploy -> Manage deployments -> Edit -> New version -> Deploy.');
}
function logAdmin_(token, id, title, org, ws, aid, discom) {
  regTab_('Admins').appendRow([Utilities.formatDate(new Date(), TZ, 'dd.MM.yyyy HH:mm'), googleEmail_(token), aid, discom, org, ws.join(', '), title, 'https://docs.google.com/spreadsheets/d/' + id + '/edit', id]);
  regAddUser_(aid, id, 'ADMIN');
  try { CacheService.getScriptCache().removeAll(['discoms', 'dsmap']); } catch (e) {}
}

/* ---------------- Division name (EWD ...) of a Division Head ---------------- */
function adminSetDivision(token, name) {
  var a = auth_(token); if (!a.admin) throw new Error('Admin only.');
  var n = String(name || '').replace(/\s+/g, ' ').trim().toUpperCase();
  if (!/^EWD/.test(n)) throw new Error('The division name must start with EWD, for example: EWD AZAMGARH.');
  if (!/^EWD[\s.\-]*[A-Z0-9][A-Z0-9 .&()\-]{0,38}$/.test(n)) throw new Error('Write the division name after EWD, for example: EWD AZAMGARH (letters, numbers, space . & ( ) - only).');
  cfgSet_('Division', n); cfgBump_();
  var sh = regTab_('Admins');
  if (!String(sh.getRange(1, 10).getValue()).trim()) sh.getRange(1, 10).setValue('Division (EWD)').setFontWeight('bold');
  regAdmins_().forEach(function (r, i) { if (String(r[8]).trim() === SS_ID) sh.getRange(i + 2, 10).setValue(n); });
  return n;
}

/* ---------------- DISCOM head (created by the app Creator) ----------------
 * Registry tab "DiscomHeads": Login ID | DISCOM | Name | Google email | One-time password | Password hash | Active | Must change | Created
 * The Creator (Google account that owns this script) adds a head with an e-mail address; a one-time password is e-mailed and also kept
 * in this tab until the head sets his own password. A DISCOM head can only READ the summary of every sheet registered under his DISCOM.   */
var DH_TAB = 'DiscomHeads', DH_HDR = ['Login ID', 'DISCOM', 'Name', 'Google email', 'One-time password', 'Password hash', 'Active', 'Must change', 'Created', 'Tabs'];
var DH_TABS = ['summary', 'comparison', 'data', 'drivers', 'kitty'], DH_TABS_DEF = ['summary', 'comparison'];
var DH_FN_TAB = { getData: 'data', getCols: 'data', getDataRange: 'data', getList: 'data', getDrivers: 'drivers', getKitty: 'kitty', getKittyPrev: 'kitty' };
function dhTabsParse_(v) {
  var a = String(v || '').toLowerCase().split(/[,;\s]+/).filter(function (x) { return DH_TABS.indexOf(x) >= 0; });
  return a.length ? DH_TABS.filter(function (t) { return a.indexOf(t) >= 0; }) : DH_TABS_DEF.slice();
}
function dhTabsClean_(arr) {
  var a = (arr instanceof Array ? arr : String(arr || '').split(',')).map(function (x) { return String(x).trim().toLowerCase(); });
  var o = DH_TABS.filter(function (t) { return a.indexOf(t) >= 0; });
  if (!o.length) throw new Error('Select at least one tab for the DISCOM head.');
  return o;
}
function dhTab_() {
  var ss = regSS_(), sh = ss.getSheetByName(DH_TAB);
  if (sh && !String(sh.getRange(1, 10).getValue()).trim()) sh.getRange(1, 10).setValue('Tabs').setFontWeight('bold');
  if (!sh) { sh = ss.insertSheet(DH_TAB); sh.getRange(1, 1, 1, DH_HDR.length).setValues([DH_HDR]).setFontWeight('bold'); sh.setFrozenRows(1); sh.getRange(2, 1, 400, DH_HDR.length).setNumberFormat('@'); }
  return sh;
}
function dhAll_() {
  var sh = dhTab_(), n = sh.getLastRow(), out = [];
  if (n < 2) return out;
  sh.getRange(2, 1, n - 1, DH_HDR.length).getValues().forEach(function (r, i) {
    if (!String(r[0]).trim()) return;
    out.push({ row: i + 2, id: String(r[0]).trim(), discom: dnorm_(r[1]), name: String(r[2] || r[0]), email: String(r[3]).trim().toLowerCase(), otp: String(r[4] || ''), hash: String(r[5] || ''), active: bool_(r[6]), must: bool_(r[7]), tabs: dhTabsParse_(r[9]) });
  });
  return out;
}
function dhFind_(id) { id = String(id || '').trim().toLowerCase(); return dhAll_().filter(function (x) { return x.id.toLowerCase() === id; })[0] || null; }
function dhMe_(h) { return { name: h.id, display: h.name, role: 'DISCOM HEAD', dh: true, canEdit: false, admin: false, must: h.must, tabs: h.tabs, scope: [], wsAll: [], org: h.discom, discom: h.discom, division: '' }; }
function dhAuth_(token, allowMust, tab) {
  var raw = token ? CacheService.getScriptCache().get('s_' + token) : null, se = null;
  if (raw) { try { se = JSON.parse(raw); } catch (e) { se = null; } }
  if (!se || !se.dh) throw new Error('SESSION');
  var h = dhFind_(se.k);
  if (!h || !h.active) throw new Error('SESSION');
  if (h.must && !allowMust) throw new Error('MUSTCHANGE');
  if (tab && h.tabs.indexOf(tab) < 0) throw new Error('The app Creator has not given you access to this tab.');
  return h;
}
function dhDivList_(h) {                                   // divisions (sheets) registered under the DISCOM of this head
  var seen = {}, out = [];
  regAdmins_().forEach(function (r) {
    var sid = String(r[8]).trim(); if (!sid || seen[sid] || dnorm_(r[3]) !== h.discom) return;
    seen[sid] = 1;
    out.push({ name: String(r[9] || '').trim().toUpperCase() || ('(division name not set) ' + (String(r[4] || r[2]).trim() || sid.slice(0, 6))), sid: sid });
  });
  return out.sort(function (a, b) { return a.name < b.name ? -1 : a.name > b.name ? 1 : 0; });
}
function dhDivs(token) {
  var h = dhAuth_(token), cur = CacheService.getScriptCache().get('dp_' + token) || '', l = dhDivList_(h);
  return { divs: l.map(function (x) { return x.name; }), cur: (l.filter(function (x) { return x.sid === cur; })[0] || {}).name || '' };
}
function dhPick(token, name) {                            // the head chooses which division the Data / Driver / Kitty tabs show
  var h = dhAuth_(token), x = dhDivList_(h).filter(function (d) { return d.name === String(name || ''); })[0];
  if (!x) throw new Error('Division not found.');
  CacheService.getScriptCache().put('dp_' + token, x.sid, 21600); return x.name;
}
function dhView_(token, se) {                             // pseudo user: sees ALL workshops of the chosen division, can change nothing
  var h = dhFind_(se.k);
  if (!h || !h.active) throw new Error('SESSION');
  if (h.must) throw new Error('MUSTCHANGE');
  var tab = DH_FN_TAB[CUR_FN];
  if (!tab) throw new Error('This is not available for a DISCOM head.');
  if (h.tabs.indexOf(tab) < 0) throw new Error('The app Creator has not given you access to this tab.');
  var sid = CacheService.getScriptCache().get('dp_' + token);
  if (!sid || !dhDivList_(h).some(function (d) { return d.sid === sid; })) throw new Error('Select a division first.');
  useSheet_(sid);
  return { name: h.id, key: 'dh', display: h.name, role: 'DISCOM HEAD', canEdit: false, admin: false, must: false, active: true, scope: ['ALL'], row: 0, pass: '' };
}
function dhSession_(h) {
  var t = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('s_' + t, JSON.stringify({ dh: 1, k: h.id }), 21600);
  return t;
}
function dhLogin_(h, password) {
  var cache = CacheService.getScriptCache(), lk = 'fail_dh_' + h.id.toLowerCase();
  if (+(cache.get(lk) || 0) >= 5) return { ok: false, msg: 'Too many wrong passwords. Try again after 10 minutes.' };
  if (!h.active || !h.hash || !chk_(h.hash, String(password || ''))) { cache.put(lk, String(+(cache.get(lk) || 0) + 1), 600); return { ok: false, msg: 'The ID or password is wrong.' }; }
  cache.remove(lk);
  return { ok: true, token: dhSession_(h), sid: '', me: dhMe_(h) };
}
function dhSetPass_(h, pw, must) {
  var sh = dhTab_();
  sh.getRange(h.row, 5).setValue(must ? h.otp : '');          // the one-time password stays visible to the Creator only until the head sets his own
  sh.getRange(h.row, 6).setValue(mk_(pw)); sh.getRange(h.row, 8).setValue(!!must);
}
function otpMake_(n) { var s = '', ch = 'ABCDEFGHJKMNPQRSTUVWXYZ23456789'; for (var i = 0; i < n; i++) s += ch.charAt(Math.floor(Math.random() * ch.length)); return s; }
function mail_(to, subject, body) {
  try { MailApp.sendEmail(to, subject, body); return true; } catch (e) { Logger.log('mail: ' + e.message); return false; }
}
// --- Creator = the Google account that owns this script (checked through Google, nothing is typed)
function creatorEmail_() {
  var e = PropertiesService.getScriptProperties().getProperty('CREATOR_EMAIL') || '';
  if (!e) { try { e = Session.getEffectiveUser().getEmail(); } catch (x) { e = ''; } }
  return String(e).trim().toLowerCase();
}
function creatorLogin(gtoken) {
  var em = googleEmail_(gtoken).toLowerCase(), owner = creatorEmail_();
  if (!em) throw new Error('Google sign-in could not be verified (' + (GERR_ || 'unknown reason') + '). Please sign in again; if it keeps failing the app Creator must run authorizeAll once in the Apps Script editor and deploy a new version.');
  if (!owner || em !== owner) throw new Error('This Google account (' + em + ') is not the Creator of this app.');
  var t = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('c_' + t, em, 7200);
  return { token: t, email: em };
}
function crAuth_(token) { if (!token || !CacheService.getScriptCache().get('c_' + token)) throw new Error('SESSION'); }
function dhRows_(list) { return list.map(function (h) { return { id: h.id, discom: h.discom, name: h.name, email: h.email, otp: h.otp, active: h.active, must: h.must, tabs: h.tabs }; }); }
function creatorHeads(token) { crAuth_(token); return { heads: dhRows_(dhAll_()), discoms: discomList() }; }
function dhIssue_(h, isNew) {                               // new one-time password -> sheet + e-mail
  var otp = otpMake_(8), sh = dhTab_();
  sh.getRange(h.row, 5).setValue(otp); sh.getRange(h.row, 6).setValue(mk_(otp)); sh.getRange(h.row, 7).setValue(true); sh.getRange(h.row, 8).setValue(true);
  var sent = mail_(h.email, 'TF-Tracker - your DISCOM head login', 'Hello ' + h.name + ',\n\nYou have been given DISCOM head access (' + h.discom + ') in TF-Tracker.\n\nLogin ID: ' + h.id + '\nOne-time password: ' + otp + '\n\nOpen the TF-Tracker home page, choose "Staff login", sign in with this ID and password. You will be asked to set your own password.\nIf you forget the password later, use "Forgot password" and sign in with this Google account (' + h.email + ').');
  return { otp: otp, mailed: sent };
}
function creatorAddHead(token, discom, name, email, id, tabs) {
  crAuth_(token); var tb = dhTabsClean_(tabs && tabs.length ? tabs : DH_TABS_DEF);
  discom = dnorm_(discom); name = String(name || '').replace(/\s+/g, ' ').trim().slice(0, 40); email = String(email || '').trim().toLowerCase();
  if (!/^[A-Z0-9][A-Z0-9 .&()\-]{1,29}$/.test(discom)) throw new Error('Enter the DISCOM name (for example PUVVNL).');
  if (!name) throw new Error('Enter the name of the DISCOM head.');
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Enter a valid Google e-mail address.');
  var l = lock_();
  try {
    var all = dhAll_(), base = String(id || '').trim() || ('DH-' + discom.replace(/[^A-Z0-9]/g, '').slice(0, 10)), nid = base, k = 1;
    if (!/^DH[A-Za-z0-9._@\-]{1,28}$/i.test(base)) throw new Error('The login ID of a DISCOM head must start with DH (for example DH-PUVVNL).');
    function taken(x) { return all.some(function (h) { return h.id.toLowerCase() === x.toLowerCase(); }) || !!regUsers_()[x.toLowerCase()]; }
    while (taken(nid)) nid = base + '-' + (++k);
    dhTab_().appendRow([nid, discom, name, email, '', '', true, true, Utilities.formatDate(new Date(), TZ, 'dd.MM.yyyy HH:mm'), tb.join(',')]);
    var h = dhFind_(nid), r = dhIssue_(h, true);
    return { id: nid, otp: r.otp, mailed: r.mailed };
  } finally { l.releaseLock(); }
}
function creatorResetHead(token, id) {
  crAuth_(token); var h = dhFind_(id); if (!h) throw new Error('Not found.');
  var r = dhIssue_(h, false); CacheService.getScriptCache().remove('fail_dh_' + h.id.toLowerCase());
  return { id: h.id, otp: r.otp, mailed: r.mailed };
}
function creatorSetTabs(token, id, tabs) {                  // the Creator chooses the tabs of one DISCOM head (Admin is never possible)
  crAuth_(token); var h = dhFind_(id); if (!h) throw new Error('Not found.');
  var tb = dhTabsClean_(tabs); dhTab_().getRange(h.row, 10).setNumberFormat('@').setValue(tb.join(',')); return tb;
}
function creatorSetActive(token, id, active) {
  crAuth_(token); var h = dhFind_(id); if (!h) throw new Error('Not found.');
  dhTab_().getRange(h.row, 7).setValue(!!active); return true;
}
// --- DISCOM head: read-only summary of every sheet registered under his DISCOM
function getDiscomSummary(token, p) {                       // result cached 2 minutes (the Heads read many sheets)
  var h = dhAuth_(token, false, 'summary'), cc = CacheService.getScriptCache(), key = 'ds_' + h.discom.replace(/[^A-Z0-9]/g, '').slice(0, 30) + '_' + [p.date, p.from, p.to].join('-'), hit = null;
  try { hit = bigGet_(cc, key); } catch (e) {}
  if (hit) { try { return JSON.parse(hit); } catch (e) {} }
  var out = discomSummaryCalc_(h, p);
  try { bigPut_(cc, key, JSON.stringify(out), 120); } catch (e) {}
  return out;
}
function discomSummaryCalc_(h, p) {
  var c = ctx_(p), D = c.D, F = c.F, T = c.T, seen = {}, divs = [], capSet = {};
  regAdmins_().forEach(function (r) {
    var sid = String(r[8]).trim(); if (!sid || seen[sid] || dnorm_(r[3]) !== h.discom) return;
    seen[sid] = 1;
    var name = String(r[9] || '').trim().toUpperCase() || ('(division name not set) ' + (String(r[4] || r[2]).trim() || sid.slice(0, 6)));
    var rows = rowsWithArc_(sid, F), wm = {}, wl = [];
    rows.forEach(function (x) {
      if (!x[I_PR]) return;
      var pr = dk_(x[I_PRD]), is = dk_(x[I_ISS]), rt = dk_(x[I_RET]), cap = capOf_(x[I_CAP]);
      var w = String(x[I_WS] || '').replace(/\s+/g, ' ').trim().toUpperCase() || '(blank)', o = wm[w];
      if (!o) { o = wm[w] = { name: w, v: [0, 0, 0, 0, 0, 0], pr: {}, pend: {} }; wl.push(w); }
      if (isDup_(x[I_FIRM])) { if (pr >= F && pr <= T) o.v[5]++; return; }              // wrong / duplicate PR: own column only
      var tillPR = pr >= F && pr <= T, pend = pr > 0 && pr <= T && (is === 0 || is > T) && (rt === 0 || rt > T);
      if (pr === D) o.v[0]++;
      if (is === D) o.v[1]++;
      if (tillPR) { o.v[2]++; o.pr[cap] = (o.pr[cap] || 0) + 1; capSet[cap] = 1; }
      if (is >= F && is <= T) o.v[3]++;
      if (pend) { o.v[4]++; o.pend[cap] = (o.pend[cap] || 0) + 1; capSet[cap] = 1; }
    });
    var d = { name: name, ws: wl.sort().map(function (w) { return wm[w]; }), v: [0, 0, 0, 0, 0, 0], pr: {}, pend: {} };
    d.ws.forEach(function (o) { addTo_(d.v, o.v); addMap_(d.pr, o.pr); addMap_(d.pend, o.pend); });
    divs.push(d);
  });
  divs.sort(function (a, b) { return a.name.localeCompare(b.name, undefined, { numeric: true }); });
  var tot = { v: [0, 0, 0, 0, 0, 0], pr: {}, pend: {} };
  divs.forEach(function (d) { addTo_(tot.v, d.v); addMap_(tot.pr, d.pr); addMap_(tot.pend, d.pend); });
  STD_CAPS.forEach(function (k) { capSet[k] = 1; });
  var caps = Object.keys(capSet).map(Number).sort(function (a, b) { return (a || 1e9) - (b || 1e9); });
  return { date: p.date, from: p.from, to: p.to, days: dayCount_(c), discom: h.discom, caps: caps, divs: divs, total: tot };
}

/* ---------------- Admin: delete all data only with a Gmail one-time password ---------------- */
function adminEmail_() {
  var em = '';
  regAdmins_().forEach(function (r) { if (String(r[8]).trim() === SS_ID && String(r[1]).trim()) em = String(r[1]).trim(); });
  return em;
}
function adminDeleteOtp(token) {
  var a = auth_(token); if (!a.admin) throw new Error('Admin only.');
  var em = adminEmail_();
  if (!em) throw new Error('Your Google e-mail is not saved in the registry. Register this sheet again from the setup screen (same ID + password), then try again.');
  var code = ('' + Math.floor(100000 + Math.random() * 900000)), c = CacheService.getScriptCache();
  c.put('del_' + SS_ID, hash_(code, SS_ID.slice(0, 8)), 600); c.remove('delf_' + SS_ID);
  if (!mail_(em, 'TF-Tracker - confirm deleting ALL data', 'Someone asked to delete ALL DATA (DATA and ARCHIVE tabs) of the TF-Tracker sheet "' + ss_().getName() + '".\n\nOne-time password: ' + code + '\n\nIt is valid for 10 minutes. If this was not you, ignore this e-mail and change your password.'))
    throw new Error('The e-mail could not be sent. Please contact the app Creator.');
  return { sentTo: em.replace(/^(.).*(@.*)$/, '$1***$2') };
}
function adminDeleteData(token, otp, password) {
  var a = auth_(token); if (!a.admin) throw new Error('Admin only.');
  var c = CacheService.getScriptCache(), fk = 'delf_' + SS_ID, tries = +(c.get(fk) || 0);
  if (tries >= 5) { c.remove('del_' + SS_ID); throw new Error('Too many wrong attempts. Ask for a new one-time password.'); }
  var want = c.get('del_' + SS_ID);
  if (!want) throw new Error('The one-time password has expired. Ask for a new one.');
  if (!chk_(a.pass, String(password || ''))) { c.put(fk, String(tries + 1), 600); throw new Error('Your password is wrong.'); }
  if (hash_(String(otp || '').trim(), SS_ID.slice(0, 8)) !== want) { c.put(fk, String(tries + 1), 600); throw new Error('The one-time password is wrong.'); }
  c.remove('del_' + SS_ID); c.remove(fk);
  var l = lock_();
  try {
    [DATA, ARC].forEach(function (n) { var sh = ss_().getSheetByName(n); if (sh && sh.getLastRow() > 1) sh.deleteRows(2, sh.getLastRow() - 1); });
    ARCH_IX = null; cfgSet_('LastDeleteAll', Utilities.formatDate(new Date(), TZ, 'dd.MM.yyyy HH:mm') + ' by ' + a.name);
  } finally { bump_(); l.releaseLock(); }
  return true;
}

/* ---------------- daily backup: ONE file per organisation, updated every day ----------------
 * A time trigger (installed by the Creator) copies DATA, ARCHIVE, Drivers, Users and Config of every registered Sheet into
 * ONE backup file per organisation (in the Creator's Drive). The same file is overwritten each day, so storage never grows.
 * Safety: if a Sheet suddenly holds far fewer records than its last backup (wiped by mistake), the old backup is NOT overwritten. */
var BK_TABS = [DATA, ARC, DRV, USERS, CONF, 'STATS', 'KittyDays', 'KittyDay', 'KittyOpen', 'KittyReq', 'KittyPrev'];
var BK_HDR = ['Admin ID', 'Organisation', 'Sheet ID (source)', 'Backup file ID', 'Backup file link', 'Last backup', 'Records (DATA+ARCHIVE)', 'Status'];
function bkTab_() {
  var sh = regTab_('Backups');
  if (sh.getLastRow() < 1) { sh.getRange(1, 1, 1, BK_HDR.length).setValues([BK_HDR]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
function bkAll_() {
  var sh = bkTab_(), n = sh.getLastRow(), out = [];
  if (n >= 2) sh.getRange(2, 1, n - 1, BK_HDR.length).getValues().forEach(function (r, i) {
    out.push({ row: i + 2, aid: String(r[0]).trim(), org: String(r[1]).trim(), sid: String(r[2]).trim(), fid: String(r[3]).trim(), url: String(r[4]).trim(), last: String(r[5]).trim(), rows: +r[6] || 0, status: String(r[7]).trim() });
  });
  return out;
}
function bkFind_(sid) { return bkAll_().filter(function (x) { return x.sid === sid; })[0] || null; }
function bkNow_() { return Utilities.formatDate(new Date(), TZ, 'dd.MM.yyyy HH:mm'); }
function bkRead_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh || sh.getLastRow() < 1 || sh.getLastColumn() < 1) return [];
  return sh.getRange(1, 1, sh.getLastRow(), sh.getLastColumn()).getValues();
}
function bkWrite_(ss, name, vals) {                       // overwrite one tab with the given rows (text format, in chunks)
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getLastRow() > 0) sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).clearContent();
  if (!vals.length) return;
  var w = 0; vals.forEach(function (r) { if (r.length > w) w = r.length; });
  vals = vals.map(function (r) { var o = r.slice(); while (o.length < w) o.push(''); return o; });
  if (sh.getMaxColumns() < w) sh.insertColumnsAfter(sh.getMaxColumns(), w - sh.getMaxColumns());
  if (sh.getMaxRows() < vals.length) sh.insertRowsAfter(sh.getMaxRows(), vals.length - sh.getMaxRows());
  for (var i = 0; i < vals.length; i += 4000) {
    var part = vals.slice(i, i + 4000);
    sh.getRange(i + 1, 1, part.length, w).setNumberFormat('@').setValues(part);
  }
}
function bkDropBlank_(ss) { try { var s = ss.getSheetByName('Sheet1'); if (s && ss.getSheets().length > 1 && s.getLastRow() === 0) ss.deleteSheet(s); } catch (e) {} }
function bkSet_(rec, vals) { bkTab_().getRange(rec.row, 1, 1, BK_HDR.length).setValues([vals]); }
// copies one organisation. Returns the status text.
function backupOne_(aid, org, sid, force) {
  var rec = bkFind_(sid), src;
  try { src = SpreadsheetApp.openById(sid); } catch (e) {
    if (rec) bkSet_(rec, [rec.aid, rec.org, rec.sid, rec.fid, rec.url, rec.last, rec.rows, 'SOURCE SHEET NOT FOUND (backup kept as it is) - ' + bkNow_()]);
    return 'SOURCE SHEET NOT FOUND';
  }
  var data = bkRead_(src, DATA), arc = bkRead_(src, ARC), cnt = 0, i;
  for (i = 1; i < data.length; i++) if (String(data[i][I_PR] == null ? '' : data[i][I_PR]).trim() !== '') cnt++;
  cnt += Math.max(arc.length - 1, 0);
  if (rec && !force && rec.rows >= 20 && cnt < rec.rows * 0.7) {
    var msg = 'HELD - the Sheet has ' + cnt + ' records but the last backup had ' + rec.rows + '. Backup NOT overwritten. ' + bkNow_();
    bkSet_(rec, [rec.aid, rec.org, rec.sid, rec.fid, rec.url, rec.last, rec.rows, msg]);
    return msg;
  }
  var bf = null;
  if (rec && rec.fid) { try { bf = SpreadsheetApp.openById(rec.fid); } catch (e) { bf = null; } }
  if (!bf) bf = SpreadsheetApp.create('TF-Tracker Backup - ' + (org || aid) + ' - ' + aid);
  bkWrite_(bf, DATA, data); bkWrite_(bf, ARC, arc);
  [DRV, USERS, CONF].forEach(function (n) { bkWrite_(bf, n, bkRead_(src, n)); });
  bkDropBlank_(bf);
  var row = [aid, org, sid, bf.getId(), bf.getUrl(), bkNow_(), cnt, 'OK'];
  if (rec) bkSet_(rec, row); else bkTab_().appendRow(row);
  return 'OK';
}
function bkSources_() {                                   // one entry per registered Sheet
  var seen = {}, out = [];
  regAdmins_().forEach(function (r) {
    var sid = String(r[8]).trim(); if (!sid || seen[sid]) return; seen[sid] = 1;
    out.push({ aid: String(r[2]).trim(), org: String(r[4]).trim(), sid: sid });
  });
  return out;
}
function bkRun_(only, force) {                            // all sheets (oldest backup first) within ~4.5 minutes, or one
  var t0 = Date.now(), list = bkSources_(), last = {}, res = [];
  bkAll_().forEach(function (x) { last[x.sid] = x.last ? x.last.slice(6, 10) + x.last.slice(3, 5) + x.last.slice(0, 2) + x.last.slice(11) : ''; });
  if (only) list = list.filter(function (x) { return x.aid.toLowerCase() === String(only).toLowerCase(); });
  list.sort(function (a, b) { return (last[a.sid] || '') < (last[b.sid] || '') ? -1 : 1; });
  for (var i = 0; i < list.length; i++) {
    if (Date.now() - t0 > 270000) { res.push({ aid: list[i].aid, status: 'SKIPPED (time) - will run in the next round' }); continue; }
    var st; try { st = backupOne_(list[i].aid, list[i].org, list[i].sid, !!force); } catch (e) { st = 'ERROR: ' + ((e && e.message) || e); }
    res.push({ aid: list[i].aid, status: st });
  }
  return res;
}
function dailyBackup() {                                  // runs from the time trigger
  var l = LockService.getScriptLock();
  if (!l.tryLock(20000)) return;
  try { bkRun_('', false); } finally { l.releaseLock(); }
}
function bkTrigger_() { return ScriptApp.getProjectTriggers().filter(function (t) { return t.getHandlerFunction() === 'dailyBackup'; }); }
function creatorBackups(token) {
  crAuth_(token);
  return { on: bkTrigger_().length > 0, list: bkAll_(), sources: bkSources_().length };
}
function creatorBackupOn(token, on) {
  crAuth_(token);
  var ex = bkTrigger_();
  if (on === false) { ex.forEach(function (t) { ScriptApp.deleteTrigger(t); }); return { on: false }; }
  if (!ex.length) ScriptApp.newTrigger('dailyBackup').timeBased().everyDays(1).atHour(2).create();
  return { on: true };
}
function creatorBackupNow(token, aid, force) {
  crAuth_(token);
  var l = lock_();
  try { return bkRun_(aid || '', !!force); } finally { l.releaseLock(); }
}
function adminBackupInfo(token) {
  var a = auth_(token); if (!a.admin) throw new Error('Admin only.');
  var r = bkFind_(SS_ID);
  return r ? { has: true, last: r.last, rows: r.rows, status: r.status } : { has: false };
}
// restore DATA / ARCHIVE / Drivers of THIS sheet from the backup (needs an e-mailed one-time password + the Admin password)
function adminRestoreOtp(token) {
  var a = auth_(token); if (!a.admin) throw new Error('Admin only.');
  var rec = bkFind_(SS_ID); if (!rec || !rec.fid) throw new Error('No backup exists yet for this sheet.');
  var em = adminEmail_();
  if (!em) throw new Error('Your Google e-mail is not saved in the registry. Register this sheet again from the setup screen (same ID + password), then try again.');
  var code = ('' + Math.floor(100000 + Math.random() * 900000)), c = CacheService.getScriptCache();
  c.put('rst_' + SS_ID, hash_(code, SS_ID.slice(0, 8)), 600); c.remove('rstf_' + SS_ID);
  if (!mail_(em, 'TF-Tracker - confirm RESTORE from backup', 'Someone asked to replace the DATA, ARCHIVE and Drivers of the TF-Tracker sheet "' + ss_().getName() + '" with the backup taken on ' + rec.last + '.\n\nOne-time password: ' + code + '\n\nIt is valid for 10 minutes. If this was not you, ignore this e-mail.'))
    throw new Error('The e-mail could not be sent. Please contact the app Creator.');
  return { sentTo: em.replace(/^(.).*(@.*)$/, '$1***$2'), last: rec.last, rows: rec.rows };
}
function adminRestore(token, otp, password) {
  var a = auth_(token); if (!a.admin) throw new Error('Admin only.');
  var c = CacheService.getScriptCache(), fk = 'rstf_' + SS_ID, tries = +(c.get(fk) || 0);
  if (tries >= 5) { c.remove('rst_' + SS_ID); throw new Error('Too many wrong attempts. Ask for a new one-time password.'); }
  var want = c.get('rst_' + SS_ID);
  if (!want) throw new Error('The one-time password has expired. Ask for a new one.');
  if (!chk_(a.pass, String(password || ''))) { c.put(fk, String(tries + 1), 600); throw new Error('Your password is wrong.'); }
  if (hash_(String(otp || '').trim(), SS_ID.slice(0, 8)) !== want) { c.put(fk, String(tries + 1), 600); throw new Error('The one-time password is wrong.'); }
  c.remove('rst_' + SS_ID); c.remove(fk);
  var rec = bkFind_(SS_ID); if (!rec || !rec.fid) throw new Error('No backup exists yet for this sheet.');
  var bf; try { bf = SpreadsheetApp.openById(rec.fid); } catch (e) { throw new Error('The backup file could not be opened. Please contact the app Creator.'); }
  var l = lock_();
  try {
    [DATA, ARC, DRV].forEach(function (n) { bkWrite_(ss_(), n, bkRead_(bf, n)); });
    ensureCols_(ss_()); ARCH_IX = null; if (conv_() && statsTab_()) statsRebuild_(); cfgSet_('LastRestore', bkNow_() + ' by ' + a.name + ' (backup of ' + rec.last + ')');
  } finally { bump_(); l.releaseLock(); }
  return { ok: true, last: rec.last, rows: rec.rows };
}
// The Google Sheet itself was deleted: Admin makes a new empty Sheet (Anyone with the link - Editor) and restores everything into it
function restoreNewSheet(link, aid, password, gtoken) {
  aid = String(aid || '').trim();
  if (!aid) throw new Error('Enter your Profit Center (Admin ID).');
  var c = CacheService.getScriptCache(), fk = 'rn_' + aid.toLowerCase(), tries = +(c.get(fk) || 0);
  if (tries >= 5) throw new Error('Too many attempts. Please try again after 10 minutes.');
  c.put(fk, String(tries + 1), 600);
  var email = googleEmail_(gtoken).toLowerCase();
  if (!email) throw new Error('Google sign-in could not be verified (' + (GERR_ || 'unknown reason') + '). Please sign in again; if it keeps failing the app Creator must run authorizeAll once in the Apps Script editor and deploy a new version.');
  var adm = regAdmins_().filter(function (r) { return String(r[2]).trim().toLowerCase() === aid.toLowerCase(); })[0];
  if (!adm) throw new Error('This Profit Center ID is not registered as an Admin.');
  if (String(adm[1]).trim().toLowerCase() !== email) throw new Error('This Google account (' + email + ') is not the one used to set up this Admin ID.');
  var oldSid = String(adm[8]).trim(), rec = bkFind_(oldSid);
  if (!rec || !rec.fid) throw new Error('No backup exists for this Admin ID.');
  var id = parseId_(link); if (!id) throw new Error('Paste the link of the NEW Google Sheet.');
  var taken = regAdmins_().some(function (r) { return String(r[8]).trim() === id && String(r[2]).trim().toLowerCase() !== aid.toLowerCase(); });
  if (taken) throw new Error('This Google Sheet already belongs to another Admin ID. Use a new empty Sheet.');
  var bf; try { bf = SpreadsheetApp.openById(rec.fid); } catch (e) { throw new Error('The backup file could not be opened. Please contact the app Creator.'); }
  var ok = false;
  bkRead_(bf, USERS).forEach(function (r, i) { if (i && String(r[0]).trim().toLowerCase() === aid.toLowerCase() && chk_(String(r[1]), String(password || ''))) ok = true; });
  if (!ok) throw new Error('The password is wrong.');
  var dst = openSafe_(id), l = lock_();
  try {
    BK_TABS.forEach(function (n) { bkWrite_(dst, n, bkRead_(bf, n)); });
    bkDropBlank_(dst);
    useSheet_(id); SS_MEM = dst; ensureCols_(dst); cfgBump_(); ARCH_IX = null; usersBump_();
    cfgSet_('LastRestore', bkNow_() + ' into a new sheet (backup of ' + rec.last + ')'); cfgBump_();
    if (id !== oldSid) {                                    // point the registry to the new sheet
      var rg = regTab_('Admins'), n = rg.getLastRow();
      if (n >= 2) { var v = rg.getRange(2, 1, n - 1, ADM_HDR.length).getValues(); v.forEach(function (r, i) { if (String(r[8]).trim() === oldSid) { rg.getRange(i + 2, 8, 1, 2).setValues([['https://docs.google.com/spreadsheets/d/' + id + '/edit', id]]); } }); }
      var ru = regTab_('Users'), m = ru.getLastRow();
      if (m >= 2) { var u = ru.getRange(2, 2, m - 1, 1).getValues(); u.forEach(function (r, i) { if (String(r[0]).trim() === oldSid) ru.getRange(i + 2, 2).setValue(id); }); }
      bkSet_(rec, [rec.aid, rec.org, id, rec.fid, rec.url, rec.last, rec.rows, 'RESTORED into a new sheet ' + bkNow_()]);
      ['reg_users', 'reg_us2', 'discoms', 'dsmap'].forEach(function (k) { try { c.remove(k); } catch (e) {} });
    }
  } finally { bump_(); l.releaseLock(); }
  c.remove(fk);
  return { ok: true, sid: id, org: cfg_().org, last: rec.last, rows: rec.rows };
}

/* ---------------- Admin: forgot password ----------------
 * The Admin signs in with the SAME Google account that was used at setup (Google verifies the e-mail);
 * if it matches the e-mail stored in the registry for this Profit Center (Admin ID), a new password is set. */
function adminForgot(aid, newPw, gtoken) {
  aid = String(aid || '').trim(); newPw = String(newPw || '');
  if (!aid) throw new Error('Enter your Profit Center (Admin ID).');
  if (newPw.length < 6) throw new Error('The new password must have at least 6 characters.');
  var c = CacheService.getScriptCache(), fk = 'fp_' + aid.toLowerCase(), tries = +(c.get(fk) || 0);
  if (tries >= 5) throw new Error('Too many attempts. Please try again after 10 minutes.');
  c.put(fk, String(tries + 1), 600);
  var email = googleEmail_(gtoken).toLowerCase();
  if (!email) throw new Error('Google sign-in could not be verified (' + (GERR_ || 'unknown reason') + '). Please sign in again; if it keeps failing the app Creator must run authorizeAll once in the Apps Script editor and deploy a new version.');
  var dh = dhFind_(aid);
  if (dh) {                                                  // DISCOM head: same Google account as the e-mail the Creator entered
    if (dh.email !== email) throw new Error('This Google account (' + email + ') is not the e-mail registered for this DISCOM head ID.');
    dhSetPass_(dh, newPw, false); dhTab_().getRange(dh.row, 7).setValue(true);
    c.remove(fk); c.remove('fail_dh_' + dh.id.toLowerCase());
    return { ok: true };
  }
  var row = null;
  regAdmins_().forEach(function (r) { if (String(r[2]).trim().toLowerCase() === aid.toLowerCase() && String(r[8]).trim()) row = r; });
  if (!row) throw new Error('This Profit Center ID is not registered as an Admin.');
  if (String(row[1]).trim().toLowerCase() !== email) throw new Error('This Google account (' + email + ') is not the one used to set up this Admin ID.');
  var sid = String(row[8]).trim(); useSheet_(sid);
  var l = lock_();
  try {
    var u = users_().filter(function (x) { return x.key === aid.toLowerCase() && x.admin; })[0];
    if (!u) throw new Error('Admin user was not found in the sheet.');
    setPass_(u, newPw, false);
    var sh = ss_().getSheetByName(USERS); sh.getRange(u.row, 6).setValue(true); usersBump_();
  } finally { l.releaseLock(); }
  c.remove(fk); c.remove('fail_' + sid + '_' + aid.toLowerCase());
  return { ok: true };
}

/* ---------------- Public: complaint status (no login) ---------------- */
function dnorm_(s) { return String(s || '').replace(/\s+/g, ' ').trim().toUpperCase(); }
function discomList() {
  var c = CacheService.getScriptCache(), s = c.get('discoms');
  if (s) { try { return JSON.parse(s); } catch (e) {} }
  var seen = {}, out = [];
  regAdmins_().forEach(function (r) { var d = dnorm_(r[3]); if (d && !seen[d]) { seen[d] = 1; out.push(d); } });
  out.sort();
  try { c.put('discoms', JSON.stringify(out), 300); } catch (e) {}
  return out;
}
function bigPut_(c, key, str, ttl) {                      // CacheService values are limited to 100 KB -> gzip + split
  var z = 'p:' + str;
  try { z = 'g:' + Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(str, 'text/plain', 'x.txt')).getBytes()); } catch (e) {}
  var n = Math.ceil(z.length / 90000); if (n > 25) return;
  var o = {}; for (var k = 0; k < n; k++) o[key + '_' + k] = z.substr(k * 90000, 90000);
  o[key] = String(n); c.putAll(o, ttl);
}
function bigGet_(c, key) {
  var n = +c.get(key); if (!n) return null;
  var ks = []; for (var k = 0; k < n; k++) ks.push(key + '_' + k);
  var g = c.getAll(ks), s = ''; for (k = 0; k < n; k++) { if (g[ks[k]] == null) return null; s += g[ks[k]]; }
  if (s.indexOf('g:') === 0) { try { return Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(s.slice(2)), 'application/x-gzip', 'x.gz')).getDataAsString(); } catch (e) { return null; } }
  return s.indexOf('p:') === 0 ? s.slice(2) : s;
}
var CS_MEM = {};                                          // warm server instance keeps the rows in memory (no cache read / unzip)
function sheetRows_(sid) {                                // all DATA rows of one organisation's sheet (cached 10 min, cleared at once when the app saves)
  try {
    useSheet_(sid);
    var c = CacheService.getScriptCache(), key = 'cs' + sid + ver_(), m = CS_MEM[sid];
    if (m && m.k === key && Date.now() - m.t < 600000) return m.rows;
    var s = bigGet_(c, key), rows = null;
    if (s) { try { rows = JSON.parse(s); } catch (e) { rows = null; } }
    if (!rows) { rows = readAll_(); try { bigPut_(c, key, JSON.stringify(rows), 600); } catch (e) {} }
    CS_MEM[sid] = { k: key, t: Date.now(), rows: rows };
    return rows;
  } catch (e) { return []; }
}
function discomSids_(discom) {                              // sheet ids registered under this DISCOM (the map is cached 5 min, cleared when a new Admin registers)
  var d = dnorm_(discom); if (!d) throw new Error('Select your DISCOM.');
  var c = CacheService.getScriptCache(), s = c.get('dsmap'), map = null;
  if (s) { try { map = JSON.parse(s); } catch (e) { map = null; } }
  if (!map) {
    map = {};
    regAdmins_().forEach(function (r) { var k = dnorm_(r[3]), id = String(r[8]).trim(); if (!k || !id) return; (map[k] = map[k] || []); if (map[k].indexOf(id) < 0) map[k].push(id); });
    try { c.put('dsmap', JSON.stringify(map), 300); } catch (e) {}
  }
  return (map[d] || []).slice();
}
/* Fast search for the home page: the browser keeps a small index (no JE / driver details) and searches it locally while typing.
 * Details of one complaint (JE / driver + mobile numbers) are fetched only when the user taps it. */
var CS_MAX = 40000;
function complainVersion(discom) {                          // tiny answer: changes when any sheet of this DISCOM changes
  var sids = discomSids_(discom);
  return sids.map(function (sid) { var n = sheetRows_(sid).length; useSheet_(sid); return sid.slice(0, 6) + ver_() + ':' + n; }).join('|');
}
function complainIndex(discom) {
  var sids = discomSids_(discom), all = [], ver = complainVersion(discom);
  var n0 = new Date(Date.now() - 31 * 86400000), from = +Utilities.formatDate(n0, TZ, 'yyyy') * 10000 + (+Utilities.formatDate(n0, TZ, 'MM')) * 100 + (+Utilities.formatDate(n0, TZ, 'dd'));
  sids.forEach(function (sid) {                              // the browser index holds only the last 31 days + every PR that is not replaced yet; older records are searched on the server
    sheetRows_(sid).forEach(function (r) {
      if (!r[I_PR]) return;
      var d = dk_(r[I_PRD]), pend = !dk_(r[16]) && !dk_(r[I_RET]) && !isDup_(r[I_FIRM]);
      if (d >= from || pend) all.push(r);
    });
  });
  if (all.length > CS_MAX) return { big: true, v: ver, n: all.length };
  all.sort(function (a, b) { return dk_(b[I_PRD]) - dk_(a[I_PRD]); });
  return { v: ver, n: all.length, rows: all.map(function (r) { return [r[7], r[8], r[9], r[4], r[2], r[3], r[1], r[0], r[5], r[12], r[13], r[16], r[17]]; }) };
}
function complainDetail(discom, complain, pr) {
  complain = String(complain || '').trim().toLowerCase(); pr = String(pr || '').trim().toLowerCase();
  if (!complain && !pr) return null;
  var sids = discomSids_(discom);
  for (var s = 0; s < sids.length; s++) {
    var rows = sheetRows_(sids[s]);
    for (var i = 0; i < rows.length; i++) {
      var r = rows[i];
      if ((complain && String(r[7]).trim().toLowerCase() === complain && (!pr || String(r[8]).trim().toLowerCase() === pr)) || (!complain && String(r[8]).trim().toLowerCase() === pr)) return r;
    }
  }
  return null;
}
var SEARCH_IDX = [7, 8, 4, 3, 2, 1, 10, 11, 14, 15, 12];  // complain no, PR no, place, feeder, substation, division, JE, JE mobile, driver, driver mobile, firm
function complainSearch(discom, q) {
  var d = dnorm_(discom); q = String(q || '').trim().toLowerCase();
  if (!d) throw new Error('Select your DISCOM.');
  if (q.length < 3) return { total: 0, rows: [] };
  if (q.length > 40) q = q.slice(0, 40);
  var out = [];
  discomSids_(d).forEach(function (sid) {
    sheetRows_(sid).forEach(function (r) {
      for (var i = 0; i < SEARCH_IDX.length; i++) {
        var v = String(r[SEARCH_IDX[i]] || '').toLowerCase(); if (!v) continue;
        if (v.indexOf(q) === 0 || (v.length > q.length && /[a-z]/.test(v) && v.split(/[\s.\-\/,()]+/).some(function (w) { return w && w.indexOf(q) === 0; }))) { out.push(r); return; }
      }
    });
  });
  out.sort(function (a, b) { return dk_(b[I_PRD]) - dk_(a[I_PRD]); });
  return { total: out.length, rows: out.slice(0, 60) };
}

// o = { org, workshops: [..], admin: { id, name, pass }, gtoken (optional one-time Google token, only used to read the email) }
function setupApp(link, o) {
  o = o || {};
  var id = parseId_(link); if (!id) throw new Error('The Google Sheet link is not valid.');
  var ws = [], seen = {};
  (o.workshops || []).forEach(function (x) {
    if (!String(x || '').trim()) return;
    var n = wsName_(x), k = nk_(n);
    if (seen[k]) throw new Error('Workshop "' + n + '" is entered twice.');
    seen[k] = 1; ws.push(n);
  });
  if (!ws.length) throw new Error('Enter at least one workshop name.');
  if (ws.length > MAX_WS) throw new Error('At most ' + MAX_WS + ' workshops at a time.');
  var a = o.admin || {}, aid = String(a.id || '').trim(), apw = String(a.pass || '');
  if (!/^[A-Za-z0-9._@\-]{3,30}$/.test(aid)) throw new Error('Profit Center (Admin ID) must be 3-30 characters (A-Z, 0-9, . _ - @), no spaces.');
  if (apw.length < 6) throw new Error('Password must be at least 6 characters.');
  var org = String(o.org || '').replace(/[\r\n\t]+/g, ' ').trim().slice(0, 60);
  var discom = dnorm_(o.discom);
  if (!/^[A-Z0-9][A-Z0-9 .&()\-]{1,29}$/.test(discom)) throw new Error('Enter your DISCOM name (2-30 characters, e.g. PUVVNL).');
  var regSid = regUsers_()[aid.toLowerCase()];
  if (regSid && regSid !== id) throw new Error('This Profit Center ID is already registered with another Google Sheet. Log in from the home page instead.');
  var ss = openSafe_(id), l = lock_();
  try {
    useSheet_(id);
    var st = state_(ss);
    if (st === 'ready') {                                    // already set up: allow RE-REGISTERING it (fixes a missing DISCOM / registry entry) if the Admin ID + password are right
      var au = users_().filter(function (x) { return x.key === aid.toLowerCase(); })[0];
      if (!au || !au.admin || !chk_(au.pass, apw)) throw new Error('This Google Sheet is already set up. To register it again, enter the same Profit Center and password that you created.');
      cfgSet_('Discom', discom); cfgBump_();
      var admRows = regAdmins_(), known = admRows.some(function (r) { return String(r[8]).trim() === id && dnorm_(r[3]) === discom; });
      if (!known) logAdmin_(o.gtoken, id, ss.getName(), cfg_().org, ws_(), aid, discom);
      else {                                                    // registry row exists: fill in the Admin's Google e-mail if it is missing (needed for "Forgot password")
        var em = googleEmail_(o.gtoken);
        if (em) admRows.forEach(function (r, i) { if (String(r[8]).trim() === id && !String(r[1]).trim()) regTab_('Admins').getRange(i + 2, 2).setValue(em); });
      }
      users_().forEach(function (x) { regAddUser_(x.name, id, x.role); });
      return { ok: true, sid: id, org: cfg_().org, workshops: ws_(), tabs: [DATA, USERS, DRV, CONF, ARC], again: true };
    }
    if (st === 'legacy') throw new Error('This Google Sheet already contains data / users. Use a new, empty Google Sheet for setup.');
    try { build_(ss, org, ws, aid, String(a.name || '').trim().slice(0, 40), apw, discom); }
    catch (e) { throw new Error('The tabs could not be created in the Sheet. Set Share -> "Anyone with the link" to the "Editor" role, then try again. (' + e.message + ')'); }
    cfgBump_(); bump_();
    logAdmin_(o.gtoken, id, ss.getName(), org, ws, aid, discom);
    return { ok: true, sid: id, org: org, workshops: ws, tabs: [DATA, USERS, DRV, CONF, ARC] };
  } finally { l.releaseLock(); }
}

/* =================================================================
 *  Admin: workshop jodna, link, sheet ki jaanch, purana data archive
 * ================================================================= */
/* ---------------- Move to a new Google Sheet when the sheet is nearly full ----------------
 * The Admin gives the link of a NEW empty Google Sheet (Anyone with the link = Editor). After the Admin's permission (password) the data of
 * the last 1 year (and every record that is still pending) is copied to the new sheet, together with users, drivers, kitty and settings.
 * The registry then points to the new sheet, so the app loads from it (fast). The old sheet is kept untouched; Summary, Comparison and the
 * "Old data" list read the finished records older than the cut from it whenever someone asks for old dates. */
function migOpen_(link) {
  var id = parseId_(link); if (!id) throw new Error('Paste the link of the new Google Sheet.');
  if (id === SS_ID) throw new Error('This is the current Google Sheet. Create a NEW empty Google Sheet and paste its link.');
  if (regAdmins_().some(function (r) { return String(r[8]).trim() === id; })) throw new Error('This Google Sheet is already used by an Admin ID. Use a new empty Sheet.');
  if ((cfg_().old || []).some(function (o) { return o.sid === id; })) throw new Error('This is an old Google Sheet of this organisation. Use a new empty Sheet.');
  var ns; try { ns = SpreadsheetApp.openById(id); } catch (e) { throw new Error('The Google Sheet could not be opened. Open its Share settings and choose "Anyone with the link" -> "Editor", then paste the link again.'); }
  try { var t = ns.insertSheet('tt_check_' + rnd_(4)); ns.deleteSheet(t); } catch (e) { throw new Error('The new Google Sheet is not editable. Share it as "Anyone with the link" -> "Editor".'); }
  if (ns.getSheets().some(function (sh) { return sh.getLastRow() > 0 && sh.getLastColumn() > 0 && !(sh.getLastRow() === 1 && sh.getLastColumn() === 1 && String(sh.getRange(1, 1).getValue()) === ''); })) throw new Error('The new Google Sheet must be empty. Create a fresh Google Sheet.');
  return { id: id, ss: ns };
}
function migSplit_() {                                   // rows that go to the new sheet / rows that stay only in the old sheet
  var cut = migCut_(), all = readAll_(false), keep = [], old = 0;
  all.forEach(function (r) { if (oldElig_(r, cut)) old++; else keep.push(r); });
  return { cut: cut, keep: keep, old: old, total: all.length, arc: (function () { var a = ss_().getSheetByName(ARC); return a ? Math.max(a.getLastRow() - 1, 0) : 0; })() };
}
function adminMigrateCheck(token, link) {
  var a = auth_(token); if (!a.admin) throw new Error('Admin only.');
  var o = migOpen_(link), sp = migSplit_();
  return { id: o.id, title: o.ss.getName(), cut: kitDs_(sp.cut), toNew: sp.keep.length, stay: sp.old + sp.arc, total: sp.total + sp.arc };
}
function adminMigrate(token, link, password) {
  var a = auth_(token); if (!a.admin) throw new Error('Admin only.');
  if (!chk_(a.pass, String(password || ''))) throw new Error('The password is wrong.');
  var o = migOpen_(link), dst = o.ss, id = o.id, oldSid = SS_ID, l = lock_();
  try {
    var sp = migSplit_(), src = ss_(), hdr = src.getSheetByName(DATA).getRange(1, 1, 1, NC).getValues()[0];
    var oldCfg = cfg_(), oldList = (oldCfg.old || []).slice();
    // 1. copy the tabs into the new sheet
    bkWrite_(dst, DATA, [hdr].concat(sp.keep));
    [USERS, DRV, CONF, ST, KD, KDS, KO, KR, KP].forEach(function (n) { var v = bkRead_(src, n); if (v.length) bkWrite_(dst, n, v); });
    if (!dst.getSheetByName(ARC)) { var as = dst.insertSheet(ARC); as.getRange(1, 1, 1, 3).setValues([['Key (Workshop|PR NO)', 'Last date (YYYYMMDD)', 'Record (packed)']]).setFontWeight('bold'); as.setFrozenRows(1); }
    [DATA, USERS, DRV, CONF, ST, KD, KDS, KO, KR, KP].forEach(function (n) { var sh = dst.getSheetByName(n); if (sh) { sh.setFrozenRows(1); sh.getRange(1, 1, 1, sh.getLastColumn() || 1).setFontWeight('bold').setBackground('#e3e9fb'); } });
    bkDropBlank_(dst);
    // 2. the new sheet becomes the live one; remember the old sheet (finished records older than the cut are read from it on demand)
    oldList.push({ sid: oldSid, url: 'https://docs.google.com/spreadsheets/d/' + oldSid + '/edit', cut: sp.cut, on: bkNow_() });
    useSheet_(id); SS_MEM = dst; cfgBump_();
    cfgSet_('OldSheets', JSON.stringify(oldList)); cfgSet_('MovedTo', ''); cfgSet_('LastMove', bkNow_() + ' from ' + oldSid);
    usersBump_(); cfgBump_(); ARCH_IX = null; OLD_MEM = null; try { CacheService.getScriptCache().remove('use_' + id); } catch (e) {}
    // 3. the old sheet is marked as replaced (stale sessions must log in again)
    useSheet_(oldSid); try { cfgSet_('MovedTo', id); cfgBump_(); } catch (e) {}
    // 4. registry
    var rg = regTab_('Admins'), n = rg.getLastRow();
    if (n >= 2) rg.getRange(2, 1, n - 1, ADM_HDR.length).getValues().forEach(function (r, i) { if (String(r[8]).trim() === oldSid) rg.getRange(i + 2, 8, 1, 2).setValues([['https://docs.google.com/spreadsheets/d/' + id + '/edit', id]]); });
    var ru = regTab_('Users'), m = ru.getLastRow();
    if (m >= 2) ru.getRange(2, 2, m - 1, 1).getValues().forEach(function (r, i) { if (String(r[0]).trim() === oldSid) ru.getRange(i + 2, 2).setValue(id); });
    try { var rec = bkFind_(oldSid); if (rec) bkSet_(rec, [rec.aid, rec.org, id, rec.fid, rec.url, rec.last, rec.rows, 'Data moved to a new sheet ' + bkNow_()]); } catch (e) {}
    ['reg_users', 'reg_us2', 'discoms', 'dsmap'].forEach(function (k) { try { CacheService.getScriptCache().remove(k); } catch (e) {} });
    useSheet_(id); SS_MEM = dst; bump_();
    return { ok: true, sid: id, moved: sp.keep.length, stay: sp.old + sp.arc, cut: kitDs_(sp.cut) };
  } finally { l.releaseLock(); }
}

/* ---------------- Workshops: rename and delete ---------------- */
function wsSave_(list) {
  var sh = ss_().getSheetByName(CONF), v = sh.getRange(1, 1, sh.getLastRow(), 1).getValues(), row = 0;
  for (var i = 0; i < v.length; i++) if (String(v[i][0]).trim() === 'Workshops') { row = i + 1; break; }
  if (!row) { row = sh.getLastRow() + 1; sh.getRange(row, 1).setValue('Workshops'); }
  sh.getRange(row, 2).setNumberFormat('@').setValue(list.join(','));
  var rg = regTab_('Admins'); regAdmins_().forEach(function (r, i) { if (String(r[8]).trim() === SS_ID) rg.getRange(i + 2, 6).setValue(list.join(', ')); });
  cfgBump_(); usersBump_(); bump_();
  try { CacheService.getScriptCache().removeAll(['discoms', 'dsmap']); } catch (e) {}
}
function wsRenameCol_(sh, col, fn) {                      // fn(cellValue) -> new value or undefined (unchanged)
  if (!sh || sh.getLastRow() < 2) return 0;
  var rg = sh.getRange(2, col, sh.getLastRow() - 1, 1), v = rg.getValues(), n = 0;
  v.forEach(function (r, i) { var x = fn(r[0]); if (x !== undefined && x !== r[0]) { v[i][0] = x; n++; } });
  if (n) rg.setNumberFormat('@').setValues(v);
  return n;
}
function adminRenameWorkshop(token, oldName, newName) {
  var a = auth_(token); if (!a.admin) throw new Error('Admin only.');
  var old = wsKey_(oldName); if (!old) throw new Error('Workshop not found.');
  var nw = wsName_(newName), cur = ws_().slice();
  if (nk_(nw) === nk_(old) && nw === old) throw new Error('The name is the same.');
  if (cur.some(function (x) { return x !== old && nk_(x) === nk_(nw); })) throw new Error('A workshop with this name already exists.');
  var l = lock_(), ss = ss_(), done = {};
  try {
    var is = function (v) { return wsKey_(v) === old; };
    done.data = wsRenameCol_(ss.getSheetByName(DATA), I_WS + 1, function (v) { return is(v) ? nw : undefined; });
    done.drv = wsRenameCol_(ss.getSheetByName(DRV), 3, function (v) { return is(v) ? nw : undefined; });
    [[KD, 2], [KO, 2], [KR, 3], [KP, 2]].forEach(function (t) { wsRenameCol_(ss.getSheetByName(t[0]), t[1], function (v) { return is(v) ? nw : undefined; }); });
    wsRenameCol_(ss.getSheetByName(USERS), 7, function (v) {
      var parts = String(v || '').split(',').map(function (x) { return x.trim(); }), ch = false;
      parts = parts.map(function (x) { if (x.toUpperCase() === old) { ch = true; return nw; } return x; });
      return ch ? parts.join(',') : undefined;
    });
    var ash = ss.getSheetByName(ARC);                                   // archived records: key and packed record
    if (ash && ash.getLastRow() > 1) {
      var rg = ash.getRange(2, 1, ash.getLastRow() - 1, 3), v = rg.getValues(), ch = 0;
      v.forEach(function (r) {
        var k = String(r[0]); if (k.indexOf('|') < 0) return;
        var ws0 = k.split('|')[0]; if (wsKey_(ws0) !== old) return;
        r[0] = nw + k.slice(ws0.length); try { var arr = JSON.parse(r[2]); arr[I_WS] = nw; r[2] = JSON.stringify(arr); } catch (e) {} ch++;
      });
      if (ch) rg.setNumberFormat('@').setValues(v);
    }
    kdsMap_(function (d) { if (!d[old]) return false; d[nw] = d[old]; delete d[old]; return true; });
    if (conv_() && statsTab_()) statsRebuild_();
    cur[cur.indexOf(old)] = nw; wsSave_(cur);                           // same position in the list, so the index stays valid
    ARCH_IX = null;
  } finally { l.releaseLock(); }
  return ws_();
}
function adminDeleteWorkshop(token, name) {
  var a = auth_(token); if (!a.admin) throw new Error('Admin only.');
  var w = wsKey_(name); if (!w) throw new Error('Workshop not found.');
  var cur = ws_().slice(); if (cur.length <= 1) throw new Error('At least one workshop must remain.');
  var l = lock_(), ss = ss_();
  try {
    var n = 0, sh = ss.getSheetByName(DATA);
    if (sh && sh.getLastRow() > 1) sh.getRange(2, I_WS + 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) { if (String(r[1]).trim() !== '' && wsKey_(r[0]) === w) n++; });
    var ash = ss.getSheetByName(ARC);
    if (ash && ash.getLastRow() > 1) ash.getRange(2, 1, ash.getLastRow() - 1, 1).getValues().forEach(function (r) { if (wsKey_(String(r[0]).split('|')[0]) === w) n++; });
    if (n) throw new Error('This workshop still has ' + n + ' records. A workshop with data cannot be deleted - rename it instead, or ask the Creator to move its data first.');
    function dropRows(nm, col) {
      var t = ss.getSheetByName(nm); if (!t || t.getLastRow() < 2) return;
      var v = t.getRange(2, col, t.getLastRow() - 1, 1).getValues();
      for (var i = v.length - 1; i >= 0; i--) if (wsKey_(v[i][0]) === w) t.deleteRow(i + 2);
    }
    kdsMap_(function (d) { if (!d[w]) return false; delete d[w]; return true; });
    dropRows(DRV, 3); dropRows(KD, 2); dropRows(KO, 2); dropRows(KR, 3); dropRows(KP, 2);
    var us = ss.getSheetByName(USERS);                                    // remove it from every user's scope
    wsRenameCol_(us, 7, function (v) {
      var parts = String(v || '').split(',').map(function (x) { return x.trim(); }).filter(Boolean);
      if (parts.map(function (x) { return x.toUpperCase(); }).indexOf(w) < 0) return undefined;
      parts = parts.filter(function (x) { return x.toUpperCase() !== w; }); return parts.length ? parts.join(',') : 'NONE';
    });
    cur.splice(cur.indexOf(w), 1); wsSave_(cur);
  } finally { l.releaseLock(); }
  return ws_();
}
function adminAddWorkshop(token, name) {
  var a = auth_(token); if (!a.admin) throw new Error('Admin only.');
  var n = wsName_(name), l = lock_();
  try {
    var cur = cfgRead_().workshops;
    if (cur.some(function (x) { return nk_(x) === nk_(n); })) throw new Error('This workshop already exists.');
    if (cur.length >= MAX_WS) throw new Error('At most ' + MAX_WS + ' workshops are allowed.');
    cur.push(n);
    var sh = ss_().getSheetByName(CONF), v = sh.getRange(1, 1, sh.getLastRow(), 1).getValues(), row = 0;
    for (var i = 0; i < v.length; i++) if (String(v[i][0]).trim() === 'Workshops') { row = i + 1; break; }
    if (!row) { row = sh.getLastRow() + 1; sh.getRange(row, 1).setValue('Workshops'); }
    sh.getRange(row, 2).setNumberFormat('@').setValue(cur.join(','));
    cfgBump_(); bump_();
    return cur;
  } finally { l.releaseLock(); }
}
function adminInfo(token) {
  var a = auth_(token); if (!a.admin) throw new Error('Admin only.');
  return { sheet: 'https://docs.google.com/spreadsheets/d/' + SS_ID + '/edit', sid: SS_ID, workshops: ws_(), conv: !!cfg_().conv, ym: cfg_().ym || '' };
}
function adminArchive(token) {
  var a = auth_(token); if (!a.admin) throw new Error('Admin only.');
  return archiveOld(true);
}
// Data aur users ki sehat jaanchne ke liye (kuch badalta nahi)
function adminHealth(token) {
  var a = auth_(token); if (!a.admin) throw new Error('Admin only.');
  var ss = ss_(), out = [], d = ss.getSheetByName(DATA), us = ss.getSheetByName(USERS);
  if (!d) out.push('The DATA tab does not exist.');
  if (!us) out.push('The Users tab does not exist.');
  if (d) {
    var rows = readAll_(true), seen = {}, noPR = 0, badD = 0, badW = 0, badC = 0, dup = 0;
    rows.forEach(function (r) {
      if (r[I_PR] === '') { noPR++; return; }
      if (!dk_(r[I_PRD])) badD++;
      if (!wsKey_(r[I_WS])) badW++;
      if (!capOf_(r[I_CAP])) badC++;
      var k = rk_(r[I_WS], r[I_PR]); if (seen[k]) dup++; seen[k] = 1;
    });
    out.push('DATA: ' + rows.length + ' rows.');
    var ash = ss.getSheetByName(ARC); out.push('ARCHIVE: ' + (ash ? Math.max(ash.getLastRow() - 1, 0) : 0) + ' old records.');
    if (noPR) out.push('- ' + noPR + ' rows have an empty PR NO (not counted in the summary).');
    if (badD) out.push('- ' + badD + ' rows have an invalid/empty PR DATE.');
    if (badW) out.push('- ' + badW + ' rows have an unrecognised Workshop.');
    if (badC) out.push('- ' + badC + ' rows have an unrecognised CAPACITY (they go under Unknown).');
    if (dup) out.push('- ' + dup + ' duplicate (Workshop + PR NO).');
  }
  if (us) {
    var list = users_();
    out.push('Users: ' + list.length + ', Admin: ' + list.filter(function (x) { return x.admin; }).length + '.');
    list.forEach(function (x) { if (x.scope.indexOf('NONE') >= 0) out.push('- ' + x.name + ' : Scope NONE - enter the workshop name in the "Scope" column of the Users sheet.'); });
  }
  out.push('Workshops: ' + ws_().join(', '));
  return out;
}
