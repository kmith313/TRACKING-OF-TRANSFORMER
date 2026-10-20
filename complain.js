/* Complaint / PR status on the home page - no login.
   Fast path: the browser keeps a small index of the chosen DISCOM (no JE / driver details) and searches it while you type.
   Details (JE / driver + mobile numbers) are fetched only when a complaint is opened.
   Fallback: if the index is not ready yet (first visit) or too big, the server searches. */
(function () {
function capNum(s){var t=String(s==null?'':s).replace(/\u00a0/g,' ').toUpperCase(),m=t.match(/(?:^|[^\d.\/])(\d+)\s*K\.?\s*V\.?\s*A/)||t.match(/(?:^|[^\d.\/])(\d+)\s*KV/);return m?m[1]+' kVA':String(s||'').trim();}
  var C = { complain: 0, pr: 1, prDate: 2, place: 3, substation: 4, feeder: 5, division: 6, workshop: 7, capacity: 8, firm: 9, issueDate: 10, replDate: 11, time: 12 };   // index row layout
  var F = { workshop: 0, division: 1, substation: 2, feeder: 3, place: 4, capacity: 5, complain: 7, pr: 8, prDate: 9, jeName: 10, jeMobile: 11, firm: 12, issueDate: 13, driverName: 14, driverMobile: 15, replDate: 16, time: 17 };   // full row layout
  var LOCAL = window.TT_LOCAL_SEARCH !== false;
  var IDX = null, loadingFor = '', shown = [], runId = 0, timer = null, lastKey = '', had = false;
  var $ = function (i) { return document.getElementById(i); };
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function isDate(v) { return /\d/.test(String(v || '')); }
  function lsGet(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }
  function lsSet(k, v) { try { localStorage.setItem(k, v); } catch (e) {} }
  function meta(t) { $('cs-meta').textContent = t; }
  function rec(a, map) { var r = {}; Object.keys(map).forEach(function (k) { r[k] = String(a[map[k]] == null ? '' : a[map[k]]).trim(); }); return r; }
  function stOf(r) { return { rep: isDate(r.replDate), iss: isDate(r.issueDate), pr: true }; }
  function chip(s) { return s.rep ? '<span class="a-chip a-chip--ok"><i></i>Replaced</span>' : s.iss ? '<span class="a-chip a-chip--go">Replacement pending</span>' : '<span class="a-chip a-chip--wait">Not issued yet</span>'; }

  /* ---------- tiny IndexedDB store for the index ---------- */
  function idb(op, key, val) {
    return new Promise(function (res) {
      try {
        var rq = indexedDB.open('tt-cs', 1);
        rq.onupgradeneeded = function () { rq.result.createObjectStore('k'); };
        rq.onerror = function () { res(null); };
        rq.onsuccess = function () {
          var db = rq.result, tx = db.transaction('k', op === 'get' ? 'readonly' : 'readwrite'), st = tx.objectStore('k'), q = op === 'get' ? st.get(key) : st.put(val, key);
          q.onsuccess = function () { res(q.result || null); }; q.onerror = function () { res(null); }; tx.oncomplete = function () { db.close(); };
        };
      } catch (e) { res(null); }
    });
  }

  /* ---------- DISCOM list ---------- */
  function loadDiscoms() {
    var sel = $('cs-discom'), cached = [];
    try { cached = JSON.parse(lsGet('tt_discoms') || '[]'); } catch (e) {}
    function fill(list) {
      var cur = lsGet('tt_discom') || '';
      sel.innerHTML = list.length ? '<option value="">- Select DISCOM -</option>' + list.map(function (d) { return '<option' + (d === cur ? ' selected' : '') + '>' + esc(d) + '</option>'; }).join('') : '<option value="">No DISCOM registered yet</option>';
      if (list.indexOf(cur) < 0) sel.value = '';
    }
    fill(cached);
    if (sel.value) prepare(sel.value);
    window.call('discomList').then(function (l) { l = l || []; lsSet('tt_discoms', JSON.stringify(l)); fill(l); if (sel.value) prepare(sel.value); else empty(); }).catch(function (e) { if (!cached.length) { sel.innerHTML = '<option value="">Could not load DISCOMs</option>'; meta(e.message); } });
    if (!sel.value) empty();
  }

  /* ---------- local index ---------- */
  function build(o, d) {
    var hay = o.rows.map(function (r) { return [r[C.complain], r[C.pr], r[C.place], r[C.feeder], r[C.substation], r[C.division]].map(function (x) { return String(x || '').toLowerCase(); }); });
    IDX = { d: d, v: o.v, rows: o.rows, hay: hay, t: o.t || Date.now() };
  }
  function prepare(d) {                                   // called when a DISCOM is chosen / restored
    if (!LOCAL || (IDX && IDX.d === d) || loadingFor === d) { if (IDX && IDX.d === d) refresh(d); return; }
    loadingFor = d;
    idb('get', 'idx|' + d).then(function (c) {
      if (c && c.rows && $('cs-discom').value === d) { build(c, d); loadingFor = ''; again(); }
      return refresh(d, !c);
    });
  }
  var busyFor = {};
  function refresh(d, force) {                            // is the saved index still current? if not, download a new one
    if (busyFor[d]) return busyFor[d];
    var pr = refresh2(d, force); busyFor[d] = pr; var done = function () { delete busyFor[d]; }; pr.then(done, done); return pr;
  }
  function refresh2(d, force) {
    var fresh = IDX && IDX.d === d && Date.now() - IDX.t < 1800000;
    var p = (force || !IDX || IDX.d !== d) ? Promise.resolve(null) : window.call('complainVersion', d).catch(function () { return IDX.v; });
    return p.then(function (v) {
      if (IDX && IDX.d === d && v && v === IDX.v && fresh) return;
      return window.call('complainIndex', d).then(function (o) {
        loadingFor = '';
        if (!o || o.big || !o.rows) { IDX = { d: d, big: true }; return; }      // too large: the server keeps searching
        o.t = Date.now(); build(o, d); idb('set', 'idx|' + d, { v: o.v, rows: o.rows, t: o.t }); again();
      });
    }).catch(function () { loadingFor = ''; });
  }
  function again() { if ($('cs-q').value.trim().length >= 3) search(true); }

  function localFind(q) {
    var out = [], n = 0, rows = IDX.rows, hay = IDX.hay, re = /[\s.\-\/,()]+/;
    for (var i = 0; i < rows.length; i++) {
      var h = hay[i], hit = false;
      for (var k = 0; k < h.length && !hit; k++) {
        var v = h[k]; if (!v) continue;
        if (v.indexOf(q) === 0) hit = true;
        else if (v.length > q.length && /[a-z]/.test(v)) { var w = v.split(re); for (var j = 0; j < w.length; j++) if (w[j] && w[j].indexOf(q) === 0) { hit = true; break; } }
      }
      if (hit) { n++; if (out.length < 60) out.push(rows[i]); }
    }
    return { total: n, rows: out };
  }

  /* ---------- results ---------- */
  function empty() {
    $('cs-res').innerHTML = '<div class="a-empty"><b>&#128269;</b>Select your DISCOM, then type your<br>complaint or PR number to see its status instantly.</div>';
  }
  function row(x, i, anim) {
    var s = stOf(x);
    return '<div class="a-row ' + (s.rep ? 'a-done' : 'a-pending') + (anim ? '' : ' noanim') + '" data-i="' + i + '" style="--i:' + Math.min(i, 8) + '" tabindex="0" role="button"><div class="a-row__main"><div class="a-row__id">' + esc(x.complain || x.pr || '-') + '</div>' +
      '<div class="a-row__place">' + esc(x.place || '-') + '</div><div class="a-row__sub">' + esc([x.substation, x.feeder].filter(Boolean).join(' • ')) + (x.capacity ? ' • ' + esc(capNum(x.capacity)) : '') + '</div></div>' +
      '<div class="a-row__side">' + chip(s) + '<span class="a-row__go">View details ›</span></div></div>';
  }
  function show(r, map, q, d, isLocal) {
    var rows = (r.rows || []).map(function (a) { return rec(a, map); }); shown = rows; shown.local = isLocal; shown.d = d;
    meta(!r.total ? 'No result found' : r.total <= rows.length ? r.total + ' result' + (r.total > 1 ? 's' : '') : r.total + ' results - showing first ' + rows.length);
    var key = q + '|' + d, anim = !had && rows.length > 0; had = rows.length > 0; lastKey = key;
    $('cs-res').innerHTML = !r.total ? '<div class="a-none">No complaint / PR matched "' + esc(q) + '" in ' + esc(d) + '.</div>' : '<div class="a-results">' + rows.map(function (x, i) { return row(x, i, anim); }).join('') + '</div>';
  }
  function search(quiet) {
    var d = $('cs-discom').value, q = $('cs-q').value.trim(), my = ++runId, out = $('cs-res');
    if (!d) { had = false; empty(); return meta('First select your DISCOM.'); }
    if (q.length < 3) { had = false; empty(); return meta(IDX && IDX.d === d && !IDX.big ? 'Fast search is ready - type at least 3 characters.' : 'Type at least 3 characters of your complaint / PR number.'); }
    var ql = q.toLowerCase();
    var older = false;
    if (IDX && IDX.d === d && !IDX.big) {                  // the local index holds the last month + pending PRs; nothing found there -> older records are searched on the server
      var lr = localFind(ql); if (lr.total) { $('cs-spin').classList.remove('on'); return show(lr, C, q, d, true); }
      older = true;
    }
    $('cs-spin').classList.add('on'); meta(older ? 'Searching older records...' : 'Searching...');
    if (!quiet && !had) out.innerHTML = '<div class="a-results"><div class="a-sk"></div><div class="a-sk"></div></div>';
    new Promise(function (ok) { setTimeout(ok, older ? 400 : 0); }).then(function () { if (my !== runId) return null; return window.call('complainSearch', d, q); }).then(function (r) {
      if (r === null) return;
      if (my !== runId) return; $('cs-spin').classList.remove('on'); show(r || { total: 0, rows: [] }, F, q, d, false);
    }).catch(function (e) { if (my !== runId) return; $('cs-spin').classList.remove('on'); out.innerHTML = ''; meta(e.message); });
  }

  /* ---------- detail sheet ---------- */
  function kv(k, v, wide, cls) { return '<div class="a-kv ' + (wide ? 'a-kv--wide ' : '') + (cls || '') + '"><div class="a-kv__k">' + k + '</div><div class="a-kv__v">' + esc(v || '—') + '</div></div>'; }
  function person(role, name, mobile) {
    var m = String(mobile || '').replace(/[^\d+]/g, '');
    return '<div class="a-person"><div class="a-avatar">' + esc(String(name || role).trim().charAt(0).toUpperCase()) + '</div><div class="a-person__txt"><div class="a-person__role">' + role + '</div><div class="a-person__name">' + esc(name || '—') + '</div></div>' +
      (m ? '<a class="a-call" href="tel:' + esc(m) + '" aria-label="Call ' + role + '"><svg viewBox="0 0 24 24"><path d="M5 4h4l2 5-2.5 1.5a11 11 0 0 0 5 5L15 13l5 2v4a2 2 0 0 1-2 2A16 16 0 0 1 3 6a2 2 0 0 1 2-2z"/></svg><span>' + esc(mobile) + '</span></a>' : '') + '</div>';
  }
  function people(r) { return person('JE', r.jeName, r.jeMobile) + person('Driver', r.driverName, r.driverMobile); }
  function detail(r, loadingPeople) {
    var s = stOf(r), steps = [{ l: 'PR Created', d: r.prDate, on: true }, { l: 'Issued', d: r.issueDate, on: s.iss }, { l: 'Replaced', d: r.replDate, on: s.rep }];
    var n = steps.filter(function (x) { return x.on; }).length, pct = n <= 1 ? 0 : (n - 1) / (steps.length - 1) * 100;
    var sh = steps.map(function (x, k) { return '<li class="a-step ' + (x.on ? 'a-on' : '') + '" style="--d:' + (0.55 + k * 0.5) + 's"><span class="a-dot">' + (x.on ? '<svg viewBox="0 0 24 24"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg>' : '') + '</span><span class="a-step__label">' + x.l + '</span><span class="a-step__date">' + (x.d ? esc(x.d) : '—') + '</span></li>'; }).join('');
    return '<div class="a-detail ' + (s.rep ? 'a-done' : 'a-pending') + '"><div class="a-detail__glow"></div>' +
      '<div class="a-detail__head"><div class="a-ids"><div class="a-ids__label">COMPLAINT NUMBER</div><div class="a-ids__val">' + esc(r.complain || r.pr || '—') + '</div></div><button class="a-close" id="cs-close" aria-label="Close">✕</button></div>' +
      '<div class="a-detail__chip">' + chip(s) + '</div>' +
      '<div class="a-grid">' + kv('PR No', r.pr) + kv('PR Date', r.prDate) + kv('Division', r.division, true) + kv('Substation', r.substation) + kv('Feeder', r.feeder) + kv('Place of Damage', r.place, true) + kv('Capacity', r.capacity ? capNum(r.capacity) : '', false, 'a-cap') + (r.firm ? kv('Issued to firm', r.firm, true) : '') + '</div>' +
      '<div class="a-people" id="cs-people">' + (loadingPeople ? '<div class="a-sk"></div><div class="a-sk"></div>' : people(r)) + '</div>' +
      '<div class="a-track ' + (s.rep ? 'a-track--ok' : '') + '"><div class="a-track__line"><div class="a-track__fill" style="--p:' + pct + '%"></div></div><ul class="a-steps">' + sh + '</ul></div>' +
      (s.rep ? '<div class="a-banner"><span class="a-banner__ring"><svg viewBox="0 0 24 24"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg></span><div><b>Transformer replaced</b><span>Replacement date: ' + esc(r.replDate) + (r.time ? ' ' + esc(r.time) : '') + '</span></div><span class="a-bolt">⚡</span></div>' : '') + '</div>';
  }
  function burst(host) {
    if (!host) return; var old = host.querySelector('.a-confetti'); if (old) old.remove();
    var box = document.createElement('div'), colors = ['#22c55e', '#facc15', '#38bdf8', '#f472b6', '#a78bfa', '#fb923c']; box.className = 'a-confetti';
    for (var i = 0; i < 30; i++) {
      var p = document.createElement('i'), ang = (Math.PI * 2 * i) / 30 + Math.random() * 0.4, dist = 80 + Math.random() * 120;
      p.style.setProperty('--x', Math.cos(ang) * dist + 'px'); p.style.setProperty('--y', Math.sin(ang) * dist - 30 + 'px'); p.style.setProperty('--r', Math.random() * 720 - 360 + 'deg');
      p.style.background = colors[i % colors.length]; p.style.animationDelay = Math.random() * 0.12 + 's'; box.appendChild(p);
    }
    host.appendChild(box); setTimeout(function () { box.remove(); }, 1700);
  }
  function open(i) {
    var r = shown[i]; if (!r) return;
    var o = $('cs-ov'), sh = $('cs-sheet'), need = shown.local;
    sh.innerHTML = detail(r, need); o.hidden = false; document.body.classList.add('a-noscroll');
    void o.offsetWidth; o.classList.add('a-open'); sh.scrollTop = 0;
    $('cs-close').onclick = close;
    var host = sh.querySelector('.a-detail');
    if (stOf(r).rep) { setTimeout(function () { burst(host); }, 1500); host.addEventListener('click', function (e) { if (!e.target.closest('a,button')) burst(host); }); }
    if (need) {                                              // JE / driver come from the server only now
      var d = shown.d;
      window.call('complainDetail', d, r.complain, r.pr).then(function (a) {
        var el = $('cs-people'); if (!el || !o.classList.contains('a-open')) return;
        el.innerHTML = a ? people(rec(a, F)) : '<div class="a-none" style="padding:8px">JE / driver details are not available.</div>'; el.classList.add('late');
      }).catch(function () { var el = $('cs-people'); if (el) el.innerHTML = '<div class="a-none" style="padding:8px">Could not load JE / driver details.</div>'; });
    }
  }
  function close() { var o = $('cs-ov'); o.classList.remove('a-open'); document.body.classList.remove('a-noscroll'); setTimeout(function () { if (!o.classList.contains('a-open')) { o.hidden = true; $('cs-sheet').innerHTML = ''; } }, 320); }

  var q = $('cs-q'); if (!q) return;
  q.addEventListener('input', function () { clearTimeout(timer); var fast = IDX && IDX.d === $('cs-discom').value && !IDX.big; timer = setTimeout(search, fast ? 0 : 350); });
  q.addEventListener('keydown', function (e) { if (e.key === 'Enter') { clearTimeout(timer); search(); } });
  $('cs-discom').addEventListener('change', function () { lsSet('tt_discom', this.value); had = false; if (this.value) prepare(this.value); search(); });
  $('cs-res').addEventListener('click', function (e) { var r = e.target.closest('.a-row'); if (r) open(+r.dataset.i); });
  $('cs-res').addEventListener('keydown', function (e) { var r = e.target.closest('.a-row'); if (r && e.key === 'Enter') open(+r.dataset.i); });
  $('cs-ov').addEventListener('click', function (e) { if (e.target === $('cs-ov')) close(); });
  document.addEventListener('keydown', function (e) { if (e.key === 'Escape' && !$('cs-ov').hidden) close(); });
  (function go() { if (window.call) loadDiscoms(); else setTimeout(go, 25); })();      // wait until the main script has defined call()
})();
