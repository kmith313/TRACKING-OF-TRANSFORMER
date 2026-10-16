/*****************************************************************
 *  TRANSFORMER TRACKING APP  -  Google Apps Script (backend)
 *  Each organisation gives its Google Sheet link to the app
 *  (Share -> "Anyone with the link" -> Editor). On setup, these tabs are created in the Sheet:
 *  DATA, Users (Admin), Drivers (Driver Master), Config, ARCHIVE.
 *****************************************************************/
var SS_ID = '';                                         // set on every request from the token / login
var TZ = 'Asia/Kolkata';
var DATA = 'DATA', USERS = 'Users', DRV = 'Drivers', ARC = 'ARCHIVE', CONF = 'Config';
var ARCH_YEARS = 2;                                     // completed data older than this many years moves to ARCHIVE
var MAX_WS = 12;                                        // maximum workshops per app

// DATA sheet columns - only up to REPLACEMENT DATE + TIME (pasted in this order)
var COLS = ['Workshop','DIVISION','SUBSTATION','FEEDER','PLACE OF DAMAGE','CAPACITY','Category of Transformer','COMPLAIN NUMBER','PR NO','PR DATE','JE Name','JE Mobile','ISSUED TO FIRM','ISSUE DATE','DRIVER NAME','DRIVER MOBILE','REPLACEMENT DATE','TIME'];
var NC = COLS.length;                                   // 18
var DATE_IDX = [9,13,16];                               // columns that hold a DD.MM.YYYY date
var I_WS=0, I_DIV=1, I_CAP=5, I_PR=8, I_PRD=9, I_ISS=13, I_REP=16, I_TIME=17, EDIT_FROM=10;
var STD_CAPS = [10,16,25,63,100,160,250,400,630];

/* ---------------- web app entry (JSON API - frontend GitHub Pages par hai) ---------------- */
// The developer deploys this script ONCE:  Deploy -> New deployment -> Web app
//   Execute as: Me   |   Who has access: Anyone
// Then put the resulting /exec URL in the frontend's config.js. Every organisation's Google Sheet
// (Share -> Anyone with the link -> Editor) is read / written through this script.
var G_ = this;
var API = {   // only these functions can be called from outside
  checkSheet: 1, setupApp: 1, login: 1, getMe: 1, logout: 1, changePassword: 1,
  adminUsers: 1, adminReset: 1, adminSetActive: 1, adminAddUser: 1, adminAddWorkshop: 1, adminInfo: 1, adminArchive: 1, adminHealth: 1,
  getData: 1, saveGrid: 1, saveRows: 1, deleteRecord: 1, saveRecord: 1,
  getSummary: 1, getList: 1, getCols: 1, getDataRange: 1,
  getDrivers: 1, saveDrivers: 1, deleteDrivers: 1
};
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function doGet() { return json_({ ok: true, app: 'Transformer Tracking API', time: new Date().toISOString() }); }
function doPost(e) {
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (!req.fn || !API[req.fn] || typeof G_[req.fn] !== 'function') return json_({ ok: false, e: 'Unknown function' });
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
  return SS_MEM || (SS_MEM = SpreadsheetApp.openById(SS_ID));
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
  var sh = ss_().getSheetByName(CONF), o = { org: '', workshops: [] };
  if (!sh || sh.getLastRow() < 1) return o;
  sh.getRange(1, 1, sh.getLastRow(), 2).getValues().forEach(function (r) {
    var k = String(r[0]).trim(), v = String(r[1] == null ? '' : r[1]).trim();
    if (k === 'Org') o.org = v;
    if (k === 'Workshops') o.workshops = v.split(',').map(function (x) { return x.replace(/\s+/g, ' ').trim().toUpperCase(); }).filter(Boolean);
  });
  return o;
}
function cfg_() {
  if (CFG_MEM) return CFG_MEM;
  var c = CacheService.getScriptCache(), k = 'cfg_' + SS_ID, s = c.get(k);
  if (s) { try { return (CFG_MEM = JSON.parse(s)); } catch (e) {} }
  var o = cfgRead_();
  if (o.workshops.length) { try { c.put(k, JSON.stringify(o), 1800); } catch (e) {} }
  return (CFG_MEM = o);
}
function cfgBump_() { CFG_MEM = null; try { CacheService.getScriptCache().remove('cfg_' + SS_ID); } catch (e) {} }
function ws_() { return cfg_().workshops; }                               // ["AZAMGARH","MAU",...]
function wsAt_(i) { return i >= 0 ? (ws_()[i] || '') : ''; }
// any written workshop name -> the canonical name from the list ('' if not recognised)
function wsKey_(s) {
  var t = nk_(s); if (!t) return '';
  var W = ws_(), i;
  for (i = 0; i < W.length; i++) if (nk_(W[i]) === t) return W[i];
  var hit = [];                                                           // one contained in the other (e.g. "EWC AZAMGARH") - only when exactly one matches
  for (i = 0; i < W.length; i++) { var w = nk_(W[i]); if (w.length >= 3 && (t.indexOf(w) >= 0 || (t.length >= 3 && w.indexOf(t) >= 0))) hit.push(W[i]); }
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
  var m = t.toUpperCase().match(/(?:^|[^\d.\/])(\d+)\s*KV/);
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
  var m = s.match(/^(\d{1,2})[.\/\-](\d{1,2})[.\/\-](\d{2,4})/);
  if (m) { var y = +m[3]; if (y < 100) y += 2000; return p2_(m[1]) + '.' + p2_(m[2]) + '.' + y; }
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
function lock_() { var l = LockService.getScriptLock(); l.waitLock(25000); return l; }

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
function users_() {
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
function me_(u) {
  return { name: u.name, display: u.display, role: u.role, canEdit: u.canEdit, admin: u.admin, must: u.must, scope: u.scope, wsAll: ws_(), org: cfg_().org };
}
function setPass_(u, pw, must) {
  var sh = ss_().getSheetByName(USERS);
  sh.getRange(u.row, 2).setValue(mk_(pw));
  sh.getRange(u.row, 9).setValue(!!must);
}
function auth_(token, allowMust) {
  var raw = token ? CacheService.getScriptCache().get('s_' + token) : null, se = null;
  if (raw) { try { se = JSON.parse(raw); } catch (e) { se = null; } }
  if (!se || !se.s || !se.k) throw new Error('SESSION');
  useSheet_(se.s);
  var u = users_().filter(function (x) { return x.key === se.k; })[0];
  if (!u || !u.active) throw new Error('SESSION');
  if (u.must && !allowMust) throw new Error('MUSTCHANGE');      // the temporary password must be replaced first
  return u;
}
function login(sid, username, password) {
  sid = parseId_(sid);
  if (!sid) return { ok: false, msg: 'Please connect a Google Sheet first.' };
  username = String(username || '').trim().toLowerCase();
  var cache = CacheService.getScriptCache(), lk = 'fail_' + sid + '_' + username.slice(0, 40);
  if (+(cache.get(lk) || 0) >= 5) return { ok: false, msg: 'Too many wrong passwords. Try again after 10 minutes.' };
  try { useSheet_(sid); if (!ss_().getSheetByName(USERS)) return { ok: false, msg: 'This Google Sheet has not been set up. Ask the Admin to set it up.' }; }
  catch (e) { return { ok: false, msg: NOSHEET }; }
  var u = users_().filter(function (x) { return x.key === username; })[0];
  if (!u || !u.active || !chk_(u.pass, String(password || ''))) {
    cache.put(lk, String(+(cache.get(lk) || 0) + 1), 600);
    return { ok: false, msg: 'The ID or password is wrong.' };
  }
  cache.remove(lk);
  if (u.pass.indexOf('h1$') !== 0) setPass_(u, String(password), u.must);
  var t = Utilities.getUuid() + Utilities.getUuid();
  cache.put('s_' + t, JSON.stringify({ k: u.key, s: sid }), 21600);
  return { ok: true, token: t, me: me_(u) };
}
function getMe(token) { return me_(auth_(token, true)); }
function logout(token) { CacheService.getScriptCache().remove('s_' + token); return true; }
function changePassword(token, oldPw, newPw) {
  var u = auth_(token, true);
  if (!chk_(u.pass, String(oldPw || ''))) throw new Error('The old password is wrong.');
  newPw = String(newPw || '');
  if (newPw.length < 6) throw new Error('The new password must be at least 6 characters.');
  setPass_(u, newPw, false);
  return true;
}
function adminUsers(token) {
  var u = auth_(token); if (!u.admin) throw new Error('Admin only.');
  return users_().map(function (x) { return { name: x.name, display: x.display, role: x.role, scope: x.scope.join(','), canEdit: x.canEdit, active: x.active }; });
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
  ss_().getSheetByName(USERS).getRange(u.row, 6).setValue(!!active);
  return true;
}

function adminAddUser(token, name, display, role, scope, canEdit) {
  var a = auth_(token); if (!a.admin) throw new Error('Admin only.');
  name = String(name || '').trim();
  if (!/^[A-Za-z0-9._@\-]{3,30}$/.test(name)) throw new Error('User ID must be 3-30 characters (A-Z, 0-9, . _ - @), no spaces.');
  var l = lock_();
  try {
    if (users_().some(function (x) { return x.key === name.toLowerCase(); })) throw new Error('This User ID already exists.');
    scope = String(scope || '').toUpperCase().trim();
    if (['ALL'].concat(ws_()).indexOf(scope) < 0) throw new Error('Workshop/Scope is invalid.');
    var t = 'Tx@' + rnd_(6), sh = ss_().getSheetByName(USERS), r = Math.max(sh.getLastRow(), 1) + 1;
    sh.getRange(r, 1, 1, 9).setNumberFormat('@');
    sh.getRange(r, 1, 1, 9).setValues([[name, mk_(t), String(role || 'AE/JE').trim(), String(display || name).trim(), !!canEdit, true, scope, false, true]]);
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
 * Index row = [workshop(0/1/2/-1), division#, capacity, PR NO, prd, iss, rep, ret]; position = sheet row - 2. Khali PR wali row = 0.   */
var IDX_MEM = null, IDX_CH = 1000, IDX_TTL = 600;
function ver_() {
  var c = CacheService.getScriptCache(), k = 'dv_' + SS_ID, v = c.get(k);
  if (!v) { v = Utilities.getUuid().slice(0, 8); c.put(k, v, 21600); }
  return v;
}
function bump_() {                                     // DATA badla -> purana index bekaar
  IDX_MEM = null;
  try { CacheService.getScriptCache().put('dv_' + SS_ID, Utilities.getUuid().slice(0, 8), 21600); } catch (e) {}
}
function buildIdx_() {
  var sh = ss_().getSheetByName(DATA), n = sh ? sh.getLastRow() : 0, divs = [], dmap = {}, rows = [], W = ws_();
  if (n >= 2) {
    var vals = sh.getRange(2, 1, n - 1, I_REP + 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      var r = vals[i], pr = norm_(I_PR, r[I_PR]);
      if (pr === '') { rows.push(0); continue; }
      var dv = divKey_(r[I_DIV]), di = dmap[dv];
      if (di === undefined) { di = divs.length; dmap[dv] = di; divs.push(dv); }
      rows.push([W.indexOf(wsKey_(r[I_WS])), di, capOf_(norm_(I_CAP, r[I_CAP])), pr,
        dk_(norm_(I_PRD, r[I_PRD])), dk_(norm_(I_ISS, r[I_ISS])), dk_(norm_(I_REP, r[I_REP]))]);
    }
  }
  return { divs: divs, rows: rows };
}
function idx_() {
  if (IDX_MEM) return IDX_MEM;
  var ix = null, ver = '', uc = null, k, pre = '', cnt = -1;
  try { var sh0 = ss_().getSheetByName(DATA); cnt = sh0 ? Math.max(sh0.getLastRow() - 1, 0) : 0; } catch (e) {}
  try {
    ver = ver_(); uc = CacheService.getUserCache(); pre = 'ix' + SS_ID + ver;
    var m = uc.get(pre);
    if (m) {
      var meta = JSON.parse(m), keys = [];
      for (k = 0; k < meta.n; k++) keys.push(pre + '_' + k);
      var got = uc.getAll(keys), rows = [], ok = true;
      for (k = 0; k < meta.n; k++) { var s = got[keys[k]]; if (s === null || s === undefined) { ok = false; break; } rows = rows.concat(JSON.parse(s)); }
      if (ok && meta.len === cnt) ix = { divs: meta.divs, rows: rows };      // sheet ki rows ginti se mel khaye tabhi cache use karo
    }
  } catch (e) { ix = null; }
  if (!ix) {
    ix = buildIdx_();
    try {
      if (uc) {
        var nch = Math.ceil(ix.rows.length / IDX_CH), put = {}, cnt2 = 0;
        for (k = 0; k < nch; k++) {
          put[pre + '_' + k] = JSON.stringify(ix.rows.slice(k * IDX_CH, (k + 1) * IDX_CH));
          if (++cnt2 === 10) { uc.putAll(put, IDX_TTL); put = {}; cnt2 = 0; }
        }
        if (cnt2) uc.putAll(put, IDX_TTL);
        uc.put(pre, JSON.stringify({ n: nch, len: ix.rows.length, divs: ix.divs }), IDX_TTL);     // meta sabse aakhir me: adhura index kabhi use nahi hoga
      }
    } catch (e) {}
  }
  return (IDX_MEM = ix);
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
    out.push({ i: i, ws: w, div: ix.divs[r[1]], cap: r[2], pr: r[3], prd: r[4], iss: r[5], rep: r[6] });
  }
  // summary / list ki date 2 saal se purani ho to ARCHIVE ka data bhi apne aap jod do
  if (c && Math.min(c.D, c.F) < cutoff_()) archRead_(Math.min(c.D, c.F)).forEach(function (a) {
    var w = wsKey_(a[I_WS]);
    if (!inScope_(u, w)) return;
    out.push({ a: a, ws: w, div: divKey_(a[I_DIV]), cap: capOf_(a[I_CAP]), pr: a[I_PR], prd: dk_(a[I_PRD]), iss: dk_(a[I_ISS]), rep: dk_(a[I_REP]), arch: true });
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
    if (!rec[I_WS] && one) rec[I_WS] = one;                   // Excel me Workshop column na ho to apna workshop maan lo
    var why = '';
    if (!rec[I_PR]) why = 'PR NO is empty';
    else if (!dk_(rec[I_PRD])) why = 'PR DATE is invalid/empty (DD.MM.YYYY required)';
    else if (!wsKey_(rec[I_WS])) why = 'Workshop is empty/not recognised (AZAMGARH / MAU / BALLIA)';
    else if (!inScope_(u, wsKey_(rec[I_WS]))) why = 'This is not an entry of your workshop';
    if (why) { skipN++; if (skipped.length < 15) skipped.push('Line ' + (k + 1) + ' (PR ' + (rec[I_PR] || '-') + '): ' + why); continue; }
    var key = rk_(rec[I_WS], rec[I_PR]);
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
function oneWs_(u) { return (u.scope.indexOf('ALL') < 0 && u.scope.length === 1 && ws_().indexOf(u.scope[0]) >= 0) ? u.scope[0] : ''; }
// items = [{k: purani "WORKSHOP|PRNO" (nayi row ho to ''), r: [36 values]}]  -> row poori replace hoti hai (khali cell = mita hua)
function saveGrid(token, items) {
  var u = auth_(token);
  if (!u.canEdit) throw new Error('You do not have permission to change data.');
  if (!items || !items.length) return { added: 0, updated: 0, skipped: 0, notes: [], keys: [], bad: [], rows: {} };
  var l = lock_();
  try {
    var sh = ss_().getSheetByName(DATA), all, pos = {}, one = oneWs_(u);
    // thode rows (roz ka autosave): sirf WORKSHOP + PR NO columns padho, jis row ko badalna hai wahi poori padho (tez).
    // Bahut rows (bada paste): pehle jaisa poori sheet padho.
    var lazy = items.length <= 15;
    if (lazy) {
      var ks = keys_(sh); all = new Array(ks.length);
      ks.forEach(function (k, i) { if (k[1] !== '') pos[rk_(k[0], k[1])] = i; });
    } else {
      all = readAll_(true);
      all.forEach(function (r, i) { if (r[I_PR] !== '') pos[rk_(r[I_WS], r[I_PR])] = i; });
    }
    function row_(i) { if (all[i] === undefined) all[i] = readRow_(sh, i + 2); return all[i]; }
    var added = [], upd = {}, skipN = 0, notes = [], keys = [], updN = 0, rs = [], bad = [], fin = {};
    function skip(n, pr, why) { skipN++; bad.push(n); if (notes.length < 15) notes.push('Line ' + (n + 1) + ' (PR ' + (pr || '-') + '): ' + why); }
    items.forEach(function (it, n) {
      var src = it.r || [], rec = [], org = it.o && it.o.length ? it.o : null;
      for (var i = 0; i < NC; i++) rec.push(norm_(i, src[i]));
      if (!rec[I_WS] && one) rec[I_WS] = one;
      var why = '';
      if (!rec[I_PR]) why = 'PR NO is empty';
      else if (!dk_(rec[I_PRD])) why = 'PR DATE is invalid/empty (DD.MM.YYYY required)';
      else if (!wsKey_(rec[I_WS])) why = 'Workshop is empty/not recognised (AZAMGARH / MAU / BALLIA)';
      else if (!inScope_(u, wsKey_(rec[I_WS]))) why = 'This is not an entry of your workshop';
      else for (var d = 0; d < DATE_IDX.length; d++) { var v = rec[DATE_IDX[d]]; if (v !== '' && !dk_(v)) { why = COLS[DATE_IDX[d]] + ' date must be written as DD.MM.YYYY'; break; } }
      if (!why) MOB_IDX.forEach(function (j) {                         // mobile: sirf wahi check jo is baar badla gaya
        if (why || rec[j] === '' || (org && rec[j] === norm_(j, org[j]))) return;
        var m = mobClean_(rec[j]);
        if (m === null) why = COLS[j] + ' must be 10 digits (E+ notation is not accepted)'; else rec[j] = m;
      });
      if (why) return skip(n, rec[I_PR], why);
      var key = rk_(rec[I_WS], rec[I_PR]), old = it.k ? pos[it.k] : undefined;
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
        all[old] = mrg; upd[old] = 1; updN++;
        if (key !== it.k) { delete pos[it.k]; pos[key] = old; }
      } else if (pos[key] !== undefined) {                              // nayi row, par PR NO pehle se hai -> update (khali cell purana data nahi mitata)
        var ex = row_(pos[key]), ch = false;
        if (!canEdit_(u, ex)) return skip(n, rec[I_PR], 'You do not have permission to change this entry');
        for (var j2 = 0; j2 < NC; j2++) if (rec[j2] !== '' && rec[j2] !== ex[j2]) { ex[j2] = rec[j2]; ch = true; }
        if (ch) { upd[pos[key]] = 1; updN++; }
      } else { pos[key] = all.length; all.push(rec); added.push(rec); }
      keys.push(key); fin[n] = key;
    });
    var uk = Object.keys(upd).map(Number).sort(function (a, b) { return a - b; });
    if (uk.length > 40) writeRows_(sh, uk[0] + 2, all.slice(uk[0], uk[uk.length - 1] + 1));   // sirf badli hui rows ka hissa likho, poori sheet nahi
    else writeRuns_(sh, uk, all);                                                             // paas-paas ki rows ek saath
    if (added.length) writeRows_(sh, all.length - added.length + 2, added);
    if (rs.length) archDel_(rs);
    var rows = {}; Object.keys(fin).forEach(function (n) { var p = pos[fin[n]]; if (p !== undefined) rows[n] = row_(p); });   // sheet me jo asal me save hua
    return { added: added.length - rs.length, updated: updN, skipped: skipN, notes: notes, keys: keys, bad: bad, rows: rows };
  } finally { bump_(); l.releaseLock(); }
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
  var u = auth_(token), l = lock_();
  try {
    var sh = ss_().getSheetByName(DATA), keys = keys_(sh), ri = -1, want = rk_(ws, String(pr).trim());
    for (var i = 0; i < keys.length; i++) if (keys[i][1] !== '' && rk_(keys[i][0], keys[i][1]) === want) { ri = i + 2; break; }
    if (ri < 0) throw new Error('PR ' + pr + ' not found.');
    var row = readRow_(sh, ri);
    if (!canEdit_(u, row)) throw new Error('You do not have permission to change this entry.');
    var nv = row.slice();
    Object.keys(vals || {}).forEach(function (k) {
      var i = +k; if (i < EDIT_FROM || i >= NC) return;
      var v = norm_(i, vals[k]);
      if (DATE_IDX.indexOf(i) >= 0 && v !== '' && !dk_(v)) throw new Error(COLS[i] + ' date must be written as DD.MM.YYYY.');
      if (MOB_IDX.indexOf(i) >= 0 && v !== '' && v !== row[i]) { var mc = mobClean_(v); if (mc === null) throw new Error(COLS[i] + ' must be 10 digits (E+ notation is not accepted).'); v = mc; }
      nv[i] = v;
    });
    sh.getRange(ri, 1, 1, NC).setNumberFormat('@').setValues([nv]);
    return nv.concat([1]);
  } finally { bump_(); l.releaseLock(); }
}

/* ---------------- summary ---------------- */
function ctx_(p) {
  var c = { D: dk_(p.date), F: dk_(p.from), T: dk_(p.to) };
  if (!c.D || !c.F || !c.T) throw new Error('Enter the date in DD.MM.YYYY format.');
  if (c.F > c.T) throw new Error('The "From" date cannot be later than the "To" date.');
  return c;
}
// 11 metrics (summary ke 11 columns)
function flags_(r, c) {
  var pr = r.prd, is = r.iss, rp = r.rep, D = c.D, F = c.F, T = c.T;
  return [
    pr < F && (is === 0 || is >= F),            // 0 LAST MONTH (opening) PENDENCY
    pr === D,                                   // 1 TODAY PR
    pr === D && is > 0,                         // 2 ISSUE AGAINST TODAY'S PR
    pr === D && rp > 0,                         // 3 REPLACEMENT AGAINST TODAY'S PR
    is === D,                                   // 4 TODAY ISSUE
    rp === D,                                   // 5 TODAY REPLACEMENT
    pr >= F && pr <= T,                         // 6 TOTAL PR
    is >= F && is <= T,                         // 7 TOTAL ISSUE
    rp >= F && rp <= T,                         // 8 TOTAL REPLACEMENT
    pr <= T && (is === 0 || is > T),            // 9 TOTAL PENDING FOR ISSUE
    is > 0 && is <= T && (rp === 0 || rp > T)   // 10 TOTAL PENDING FOR REPLACEMENT
  ];
}
function zero_(n) { var a = []; for (var i = 0; i < n; i++) a.push(0); return a; }
function addTo_(a, b) { for (var i = 0; i < a.length; i++) a[i] += b[i]; return a; }
function dayCount_(c) {
  function d(k) { return new Date(Math.floor(k / 10000), Math.floor(k / 100) % 100 - 1, k % 100); }
  return Math.round((d(c.T) - d(c.F)) / 86400000) + 1;
}

function getSummary(token, p) {
  var u = auth_(token), c = ctx_(p), recs = load_(u, c), WS = ws_(), nW = WS.length;
  var capSet = {}; STD_CAPS.forEach(function (k) { capSet[k] = 1; });
  recs.forEach(function (r) { capSet[r.cap] = 1; });
  var caps = Object.keys(capSet).map(Number).sort(function (a, b) { return (a || 1e9) - (b || 1e9); });

  var wsShown = WS.filter(function (w) { return inScope_(u, w); });
  var top = {}; WS.forEach(function (w) { top[w] = zero_(11); });
  var divMap = {}, capI = {}, capR = {}, capP = {}, divCap = {}, divWs = {};
  caps.forEach(function (k) { capI[k] = zero_(nW); capR[k] = zero_(nW); capP[k] = zero_(nW); });

  recs.forEach(function (r) {
    var f = flags_(r, c), w = WS.indexOf(r.ws);
    var dv = divMap[r.div] || (divMap[r.div] = zero_(11));
    var dc = divCap[r.div] || (divCap[r.div] = {});
    if (w >= 0 && !divWs[r.div]) divWs[r.div] = r.ws;                  // ye division kis workshop ki hai
    for (var i = 0; i < 11; i++) if (f[i]) { dv[i]++; if (w >= 0) top[r.ws][i]++; }
    if (w >= 0) { if (f[9]) capI[r.cap][w]++; if (f[10]) capR[r.cap][w]++; if (f[6]) capP[r.cap][w]++; }
    if (f[6]) dc[r.cap] = (dc[r.cap] || 0) + 1;
  });

  function sumV(keys) { var s = zero_(11); keys.forEach(function (k) { if (divMap[k]) addTo_(s, divMap[k]); }); return s; }
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
  if (othersAny) push('o', 'OTHER / LIST ME NAHI', others);
  var allKeys = Object.keys(divMap);

  var topRows = wsShown.map(function (w) { return { label: w, ws: w, v: top[w] }; });
  var topTot = zero_(11); topRows.forEach(function (r) { addTo_(topTot, r.v); });

  return {
    date: p.date, from: p.from, to: p.to, days: dayCount_(c), ws: wsShown, wsAll: WS, caps: caps,
    top: { rows: topRows, total: topTot },
    div: { rows: rowsD, total: sumV(allKeys) },
    divCap: { rows: rowsC, total: sumC(allKeys) },
    capIssue: capI, capRep: capR, capPR: capP
  };
}

// summary ke kisi number par click -> poori list
function getList(token, spec) {
  var u = auth_(token), c = ctx_(spec), m = +spec.metric;
  var dv = spec.divs ? spec.divs.reduce(function (o, k) { o[k] = 1; return o; }, {}) : null;
  var cp = spec.caps ? spec.caps.reduce(function (o, k) { o[k] = 1; return o; }, {}) : null;
  for (var t = 0; t < 2; t++) {
    var out = load_(u, c).filter(function (r) {
      if (spec.ws && r.ws !== spec.ws) return false;
      if (dv && !dv[r.div]) return false;
      if (cp && !cp[r.cap]) return false;
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
  var which = basis === 'pr' ? [0] : basis === 'issue' ? [1] : basis === 'rep' ? [2] : [0, 1, 2];
  var out = [];
  var src = readAll_();
  if (F < cutoff_()) src = src.concat(archRead_(F));            // purana data ARCHIVE se
  src.forEach(function (a) {
    var wk = wsKey_(a[I_WS]);
    if (!inScope_(u, wk) || (w && wk !== w)) return;
    var d = [dk_(a[I_PRD]), dk_(a[I_ISS]), dk_(a[I_REP])], hit = false;
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
function getDrivers(token) {
  var u = auth_(token), rows = drvAll_();
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
// Sirf wahi record archive hote hain jinka issue + replacement ho chuka ho aur jinki koi bhi date 2 saal se nayi na ho.
function cutoff_() {
  var n = new Date();
  return (+Utilities.formatDate(n, TZ, 'yyyy') - ARCH_YEARS) * 10000 + (+Utilities.formatDate(n, TZ, 'MM')) * 100 + (+Utilities.formatDate(n, TZ, 'dd'));
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
  var sh = ss_().getSheetByName(ARC), n = sh ? sh.getLastRow() : 0;
  if (n < 2) return [];
  var b = sh.getRange(2, 2, n - 1, 1).getValues(), st = -1;
  for (var i = 0; i < b.length; i++) if (+b[i][0] >= minK) { st = i; break; }
  if (st < 0) return [];
  var out = [];
  sh.getRange(st + 2, 3, n - 1 - st, 1).getValues().forEach(function (r) { try { out.push(unpack_(r[0])); } catch (e) {} });
  return out;
}
// kisi purani PR par naya data paste/save ho to record ARCHIVE se nikal kar DATA me wapas aa jata hai
function restore_(key, all, pos, added, rs) {
  var rn = archIdx_()[key]; if (!rn) return;
  var rec; try { rec = unpack_(archSheet_().getRange(rn, 3).getValue()); } catch (e) { return; }
  pos[key] = all.length; all.push(rec); added.push(rec); rs.push(rn); delete ARCH_IX[key];
}
function archDel_(rowNums) {
  var sh = archSheet_(), nums = rowNums.slice().sort(function (a, b) { return b - a; });
  nums.forEach(function (n) { sh.deleteRow(n); });
  ARCH_IX = null;
}
function archiveOld(silent) {
  silent = silent === true;
  var l = lock_();
  try {
    var sh = ss_().getSheetByName(DATA); if (!sh) { if (!silent) alert_('The DATA tab does not exist -> run Setup.'); return 0; }
    var cut = cutoff_(), all = readAll_(true), keep = [], move = [], ix = archIdx_();
    all.forEach(function (r) {
      if (r[I_PR] === '') { keep.push(r); return; }
      var mx = maxDate_(r);
      if (mx > 0 && mx < cut && dk_(r[I_PRD]) && dk_(r[I_ISS]) && dk_(r[I_REP]) && wsKey_(r[I_WS]) && !ix[rk_(r[I_WS], r[I_PR])]) move.push(r);
      else keep.push(r);
    });
    if (!move.length) { if (!silent) alert_('No old data eligible for archiving was found.\n(older than 2 years, with both issue and replacement done)'); return 0; }
    var ash = archSheet_(), start = Math.max(ash.getLastRow(), 1) + 1;
    var out = move.map(function (r) { return [rk_(r[I_WS], r[I_PR]), String(maxDate_(r)), JSON.stringify(r)]; });
    var need = start + out.length - 1;
    if (need > ash.getMaxRows()) ash.insertRowsAfter(ash.getMaxRows(), need - ash.getMaxRows());
    ash.getRange(start, 1, out.length, 3).setNumberFormat('@').setValues(out);
    SpreadsheetApp.flush();
    if (ash.getLastRow() < need) throw new Error('Writing the archive did not complete - nothing in DATA was changed.');
    ash.getRange(2, 1, ash.getLastRow() - 1, 3).sort({ column: 2, ascending: true });
    // pehle bache hue records upar likho, phir neeche ka bacha hissa saaf karo (beech me fail ho to bhi data khona nahi)
    writeRows_(sh, 2, keep);
    if (all.length > keep.length) sh.getRange(keep.length + 2, 1, all.length - keep.length, NC).clearContent();
    ARCH_IX = null;
    bump_();
    compactSheets(true);
    if (!silent) alert_(move.length + ' old records were moved to ARCHIVE (cells: ' + move.length * NC + ' -> ' + move.length * 3 + ').\nThey will be included automatically for old dates in Summary / Excel download.');
    return move.length;
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
function build_(ss, org, ws, aid, adisp, apw) {
  var dt = mkTab_(ss, DATA, COLS);
  dt.getRange(1, 1, dt.getMaxRows(), NC).setNumberFormat('@');           // dates text (DD.MM.YYYY) hi rahein
  dt.setColumnWidths(1, NC, 120);
  var us = mkTab_(ss, USERS, ['User ID', 'Password', 'Role', 'Display Name', 'Can Edit', 'Active', 'Scope', 'Admin', 'Must change']);
  us.getRange(2, 1, 1, 9).setNumberFormat('@');
  us.getRange(2, 1, 1, 9).setValues([[aid, mk_(apw), 'ADMIN', adisp || aid, true, true, 'ALL', true, false]]);
  var dr = mkTab_(ss, DRV, ['ID', 'Firm Name', 'Workshop', 'Driver Name', 'Mobile No.', 'Vehicle GPS URL', 'Vehicle No.']);
  dr.getRange(2, 1, dr.getMaxRows() - 1, 7).setNumberFormat('@');
  mkTab_(ss, ARC, ['Key (Workshop|PR NO)', 'Last date (YYYYMMDD)', 'Record (packed)']);
  var cf = mkTab_(ss, CONF, ['Setting', 'Value']);
  cf.getRange(2, 1, 6, 2).setNumberFormat('@');
  cf.getRange(2, 1, 6, 2).setValues([
    ['App', 'Transformer Tracking'], ['Org', org || ''], ['Workshops', ws.join(',')], ['Version', '2'],
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

/* ---------------- Admin registry (developer only) ----------------
 * Every time someone sets up the app, one row is added to a Google Sheet in the DEVELOPER's Drive
 * ("Transformer Tracking - Admin Registry"): time, verified Google email, sheet link, workshops, admin ID.
 * The email is verified by Google (the one-time access token is checked), it is not typed by the user.
 * To find the sheet: open script.google.com -> run registryUrl() once -> see the link in the Execution log. */
function registrySheet_() {
  var props = PropertiesService.getScriptProperties(), id = props.getProperty('REGISTRY_ID'), ss = null;
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    ss = SpreadsheetApp.create('Transformer Tracking - Admin Registry');
    props.setProperty('REGISTRY_ID', ss.getId());
    var s0 = ss.getSheets()[0]; s0.setName('Admins');
    s0.getRange(1, 1, 1, 7).setValues([['Time', 'Admin Google email', 'Admin ID', 'Organisation', 'Workshops', 'Sheet title', 'Sheet link']]).setFontWeight('bold');
    s0.setFrozenRows(1);
  }
  return ss.getSheets()[0];
}
function registryUrl() { var u = registrySheet_().getParent().getUrl(); Logger.log(u); return u; }
function googleEmail_(token) {
  if (!token) return '';
  try {
    var r = UrlFetchApp.fetch('https://www.googleapis.com/oauth2/v3/userinfo', { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true });
    return r.getResponseCode() === 200 ? String(JSON.parse(r.getContentText()).email || '') : '';
  } catch (e) { return ''; }
}
function logAdmin_(token, id, title, org, ws, aid) {
  try {
    registrySheet_().appendRow([Utilities.formatDate(new Date(), TZ, 'dd.MM.yyyy HH:mm'), googleEmail_(token), aid, org, ws.join(', '), title, 'https://docs.google.com/spreadsheets/d/' + id + '/edit']);
  } catch (e) { Logger.log('registry: ' + e.message); }       // never block a setup because of the registry
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
  if (!/^[A-Za-z0-9._@\-]{3,30}$/.test(aid)) throw new Error('Admin ID must be 3-30 characters (A-Z, 0-9, . _ - @), no spaces.');
  if (apw.length < 6) throw new Error('Password must be at least 6 characters.');
  var org = String(o.org || '').replace(/[\r\n\t]+/g, ' ').trim().slice(0, 60);
  var ss = openSafe_(id), l = lock_();
  try {
    useSheet_(id);
    var st = state_(ss);
    if (st === 'ready') throw new Error('This Google Sheet is already set up. Connect with the "User" option, or log in with the Admin ID.');
    if (st === 'legacy') throw new Error('This Google Sheet already contains data / users. Use a new, empty Google Sheet for setup.');
    try { build_(ss, org, ws, aid, String(a.name || '').trim().slice(0, 40), apw); }
    catch (e) { throw new Error('The tabs could not be created in the Sheet. Set Share -> "Anyone with the link" to the "Editor" role, then try again. (' + e.message + ')'); }
    cfgBump_(); bump_();
    logAdmin_(o.gtoken, id, ss.getName(), org, ws, aid);
    return { ok: true, sid: id, org: org, workshops: ws, tabs: [DATA, USERS, DRV, CONF, ARC] };
  } finally { l.releaseLock(); }
}

/* =================================================================
 *  Admin: workshop jodna, link, sheet ki jaanch, purana data archive
 * ================================================================= */
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
  return { sheet: 'https://docs.google.com/spreadsheets/d/' + SS_ID + '/edit', sid: SS_ID, workshops: ws_() };
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
