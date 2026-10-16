/* Complaint / PR Status - shown on the login page, no login needed.
   Data: from the DATA tab of the Google Sheet the Admin connected (GViz, sheet shared "Anyone with the link").
   Only the needed columns; status goes up to REPLACEMENT DATE + TIME. */
(function () {
  var REFRESH_MS = 5 * 60 * 1000, GAP_MS = 20 * 1000, LIMIT = 60;
  var FIELDS = [
    ['workshop', ['WORKSHOP'], 0], ['division', ['DIVISION'], 1], ['substation', ['SUBSTATION'], 2], ['feeder', ['FEEDER'], 3],
    ['place', ['PLACE OF DAMAGE'], 4], ['capacity', ['CAPACITY'], 5], ['complain', ['COMPLAIN NUMBER'], 7], ['pr', ['PR NO'], 8],
    ['prDate', ['PR DATE'], 9], ['jeName', ['JE NAME'], 10], ['jeMobile', ['JE MOBILE'], 11], ['firm', ['ISSUED TO FIRM'], 12],
    ['issueDate', ['ISSUE DATE'], 13], ['driverName', ['DRIVER NAME'], 14], ['driverMobile', ['DRIVER MOBILE'], 15],
    ['replDate', ['REPLACEMENT DATE'], 16], ['time', ['TIME'], 17]
  ];
  var sid = '', records = [], updatedAt = 0, active = false, timer = null, queue = Promise.resolve(), runId = 0;
  var $ = function (i) { return document.getElementById(i); };
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function norm(s) { return String(s || '').replace(/\s+/g, ' ').trim().toUpperCase(); }
  function isDate(v) { return /\d/.test(String(v || '')); }
  function ckey() { return 'cps:' + sid; }
  function setStatus(t) { $('cs-status').textContent = t; }

  function loadCache() {
    records = []; updatedAt = 0;
    try { var c = JSON.parse(localStorage.getItem(ckey()) || 'null'); if (c && c.records) { records = c.records; updatedAt = c.t || 0; } } catch (e) {}
  }
  function parseTable(t) {
    if (!t) return [];
    var cols = (t.cols || []).map(function (c) { return norm(c && c.label); });
    var rows = (t.rows || []).map(function (r) { return (r.c || []).map(function (x) { return !x ? '' : x.f != null ? String(x.f) : x.v != null ? String(x.v) : ''; }); });
    var hdr = cols;
    if (cols.indexOf('PR NO') < 0 && rows.length) { var f = rows[0].map(norm); if (f.indexOf('PR NO') >= 0) { hdr = f; rows = rows.slice(1); } }
    var map = FIELDS.map(function (f) { var i = -1; for (var k = 0; k < f[1].length && i < 0; k++) i = hdr.indexOf(f[1][k]); return i >= 0 ? i : f[2]; });
    var out = [];
    rows.forEach(function (r) {
      var rec = {}; FIELDS.forEach(function (f, i) { rec[f[0]] = String(r[map[i]] == null ? '' : r[map[i]]).trim(); });
      if (rec.pr || rec.complain) out.push(rec);
    });
    return out;
  }
  function gviz(url) {
    return new Promise(function (res, rej) {
      var s = document.createElement('script'), g = (window.google = window.google || {});
      g.visualization = g.visualization || {}; g.visualization.Query = g.visualization.Query || {};
      var old = g.visualization.Query.setResponse, tm = setTimeout(function () { done(rej, new Error('Timeout')); }, 15000);
      function done(fn, v) { clearTimeout(tm); try { s.remove(); } catch (e) {} g.visualization.Query.setResponse = old; fn(v); }
      g.visualization.Query.setResponse = function (r) { done(res, r); };
      s.onerror = function () { done(rej, new Error('Could not load the sheet (sharing must be: Anyone with the link)')); };
      s.src = url; document.body.appendChild(s);
    });
  }
  function fetchSheet() {
    if (!sid) return Promise.resolve();
    var url = 'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(sid) + '/gviz/tq?sheet=DATA&headers=1&tqx=out:json';
    return gviz(url).then(function (r) {
      if (r && r.status === 'error') throw new Error((r.errors && r.errors[0] && r.errors[0].detailed_message) || 'Sheet error');
      records = parseTable(r.table); updatedAt = Date.now();
      try { localStorage.setItem(ckey(), JSON.stringify({ records: records, t: updatedAt })); } catch (e) {}
    });
  }
  function enqueue(fn) { return (queue = queue.then(fn, fn)); }
  function refresh(show) {
    if (show) setStatus('Loading data...');
    return enqueue(fetchSheet).then(function () { stamp(); if (show) setStatus('Updated ✓'); search(true); }, function (e) { if (show) setStatus('Refresh failed: ' + e.message); });
  }
  function stamp() { $('cs-upd').textContent = updatedAt ? 'Updated ' + new Date(updatedAt).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }) : ''; }

  function match(r, q) {
    var hay = [r.complain, r.pr, r.place, r.feeder, r.substation, r.division, r.jeName, r.jeMobile, r.driverName, r.driverMobile, r.firm];
    for (var i = 0; i < hay.length; i++) {
      var s = (hay[i] || '').toLowerCase(); if (!s) continue;
      if (s.indexOf(q) === 0) return true;
      if (s.length > q.length && /[a-z]/.test(s) && s.split(/[\s.\-\/,()]+/).some(function (w) { return w && w.indexOf(q) === 0; })) return true;
    }
    return false;
  }
  function st(r) { return { rep: isDate(r.replDate), iss: isDate(r.issueDate) }; }
  function chip(s) { return s.rep ? '<span class="cchip ok">Replaced</span>' : s.iss ? '<span class="cchip go">Replacement pending</span>' : '<span class="cchip wait">Not issued yet</span>'; }
  var shown = [];
  function search(silent) {
    var q = $('cs-q').value.trim().toLowerCase(), my = ++runId, out = $('cs-res');
    if (!q) { out.innerHTML = ''; $('cs-meta').textContent = sid ? 'Type a complaint / PR number...' : ''; return; }
    function go() {
      if (my !== runId) return;
      var m = records.filter(function (r) { return match(r, q); });
      shown = m.slice(0, LIMIT);
      $('cs-meta').textContent = !m.length ? 'No results found' : m.length <= LIMIT ? m.length + ' result' + (m.length > 1 ? 's' : '') : m.length + ' results - showing first ' + LIMIT;
      out.innerHTML = !m.length ? '<div class="cnone">No complaint / PR matched "' + esc(q) + '".</div>' : shown.map(function (r, i) {
        var s = st(r);
        return '<div class="crow ' + (s.rep ? 'done' : '') + '" data-i="' + i + '" tabindex="0" role="button"><div><b>' + esc(r.complain || r.pr) + '</b>' + (r.complain && r.pr ? ' <span class="muted">PR ' + esc(r.pr) + '</span>' : '') +
          '<div>' + esc(r.place || '-') + '</div><div class="muted">' + esc([r.substation, r.feeder].filter(Boolean).join(' • ')) + (r.capacity ? ' • ' + esc(r.capacity) + ' kVA' : '') + '</div></div><div>' + chip(s) + '</div></div>';
      }).join('');
    }
    var age = updatedAt ? Date.now() - updatedAt : Infinity;
    if (!silent && age > GAP_MS) { setStatus('Checking for fresh data...'); enqueue(fetchSheet).catch(function () {}).then(function () { stamp(); setStatus(''); go(); }); if (records.length) go(); }
    else go();
  }
  function kv(k, v, w) { return '<div class="ckv' + (w ? ' w' : '') + '"><div class="k">' + k + '</div><div class="v">' + esc(v || '-') + '</div></div>'; }
  function person(role, name, mob) {
    var m = String(mob || '').replace(/[^\d+]/g, '');
    return '<div class="cper"><div><div class="k">' + role + '</div><div class="v">' + esc(name || '-') + '</div></div>' + (m ? '<a class="btn sm" href="tel:' + esc(m) + '">☎ ' + esc(mob) + '</a>' : '') + '</div>';
  }
  function detail(r) {
    var s = st(r), steps = [['PR Created', r.prDate, true], ['Issued', r.issueDate, s.iss], ['Replaced', r.replDate + (r.time ? ' ' + r.time : ''), s.rep]];
    return '<div class="md"><div class="bar" style="justify-content:space-between;align-items:center"><div><div class="k">COMPLAINT NUMBER</div><h2 style="margin:0">' + esc(r.complain || '-') + '</h2></div><button class="btn sec" id="cs-close">✕</button></div>' +
      '<div style="margin:8px 0">' + chip(s) + '</div><div class="cgrid">' + kv('PR No', r.pr) + kv('PR Date', r.prDate) + kv('Workshop', r.workshop) + kv('Division', r.division) + kv('Substation', r.substation) + kv('Feeder', r.feeder) +
      kv('Place of Damage', r.place, true) + kv('Capacity', r.capacity ? r.capacity + ' kVA' : '') + kv('Issued to firm', r.firm) + '</div>' + person('JE', r.jeName, r.jeMobile) + person('Driver', r.driverName, r.driverMobile) +
      '<ul class="csteps">' + steps.map(function (x) { return '<li class="' + (x[2] ? 'on' : '') + '"><i>' + (x[2] ? '✓' : '') + '</i><span>' + x[0] + '</span><em>' + esc(String(x[1] || '').trim() || '-') + '</em></li>'; }).join('') + '</ul>' +
      (s.rep ? '<div class="cdone">⚡ Transformer replaced - ' + esc(r.replDate) + (r.time ? ' ' + esc(r.time) : '') + '</div>' : '') + '</div>';
  }
  function open(i) { var o = $('cs-ov'); o.innerHTML = detail(shown[i]); o.classList.remove('hide'); $('cs-close').onclick = close; }
  function close() { $('cs-ov').classList.add('hide'); }

  function init() {
    var q = $('cs-q'), t; if (!q) return;
    q.addEventListener('input', function () { clearTimeout(t); t = setTimeout(function () { search(false); }, 250); });
    $('cs-res').addEventListener('click', function (e) { var r = e.target.closest('.crow'); if (r) open(+r.dataset.i); });
    $('cs-res').addEventListener('keydown', function (e) { var r = e.target.closest('.crow'); if (r && e.key === 'Enter') open(+r.dataset.i); });
    $('cs-ov').addEventListener('click', function (e) { if (e.target === $('cs-ov')) close(); });
    $('cs-refresh').onclick = function () { refresh(true); };
  }
  window.Complain = {
    setSheet: function (id) { if (id === sid) return; sid = id || ''; loadCache(); if ($('cs-q')) { stamp(); search(true); } if (active && sid) refresh(false); },
    screen: function (name) {                       // auto-refresh only while the login screen is visible
      active = name === 'login'; clearInterval(timer);
      if (active && sid) { stamp(); refresh(false); timer = setInterval(function () { refresh(false); }, REFRESH_MS); }
    }
  };
  init();
})();
