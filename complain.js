/* Complaint / PR status on the home page - no login.
   The user picks a DISCOM; the server searches the data of every organisation (Admin sheet) registered under that DISCOM. */
(function () {
  var COL = { workshop: 0, division: 1, substation: 2, feeder: 3, place: 4, capacity: 5, complain: 7, pr: 8, prDate: 9, jeName: 10, jeMobile: 11, firm: 12, issueDate: 13, driverName: 14, driverMobile: 15, replDate: 16, time: 17 };
  var shown = [], runId = 0, timer = null;
  var $ = function (i) { return document.getElementById(i); };
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function isDate(v) { return /\d/.test(String(v || '')); }
  function lsGet(k) { try { return localStorage.getItem(k); } catch (e) { return null; } }
  function lsSet(k, v) { try { localStorage.setItem(k, v); } catch (e) {} }
  function rec(a) { var r = {}; Object.keys(COL).forEach(function (k) { r[k] = String(a[COL[k]] == null ? '' : a[COL[k]]).trim(); }); return r; }
  function st(r) { return { rep: isDate(r.replDate), iss: isDate(r.issueDate) }; }
  function chip(s) { return s.rep ? '<span class="cchip ok">Replaced</span>' : s.iss ? '<span class="cchip go">Replacement pending</span>' : '<span class="cchip wait">Not issued yet</span>'; }
  function meta(t) { $('cs-meta').textContent = t; }

  function loadDiscoms() {
    var sel = $('cs-discom'), cached = [];
    try { cached = JSON.parse(lsGet('tt_discoms') || '[]'); } catch (e) {}
    function fill(list) {
      var cur = lsGet('tt_discom') || '';
      sel.innerHTML = list.length ? '<option value="">- Select DISCOM -</option>' + list.map(function (d) { return '<option' + (d === cur ? ' selected' : '') + '>' + esc(d) + '</option>'; }).join('') : '<option value="">No DISCOM registered yet</option>';
      if (list.indexOf(cur) < 0) sel.value = '';
    }
    fill(cached);
    window.call('discomList').then(function (l) { l = l || []; lsSet('tt_discoms', JSON.stringify(l)); fill(l); }).catch(function (e) { if (!cached.length) { sel.innerHTML = '<option value="">Could not load DISCOMs</option>'; meta(e.message); } });
  }

  function search() {
    var d = $('cs-discom').value, q = $('cs-q').value.trim(), my = ++runId, out = $('cs-res');
    if (!d) { out.innerHTML = ''; return meta('First select your DISCOM.'); }
    if (q.length < 3) { out.innerHTML = ''; return meta('Type at least 3 characters of your complaint / PR number.'); }
    $('cs-spin').classList.add('on'); meta('Searching...');
    window.call('complainSearch', d, q).then(function (r) {
      if (my !== runId) return;
      $('cs-spin').classList.remove('on');
      var rows = (r && r.rows || []).map(rec); shown = rows;
      meta(!r.total ? 'No result found' : r.total <= rows.length ? r.total + ' result' + (r.total > 1 ? 's' : '') : r.total + ' results - showing first ' + rows.length);
      out.innerHTML = !r.total ? '<div class="cnone">No complaint / PR matched "' + esc(q) + '" in ' + esc(d) + '.</div>' : rows.map(function (x, i) {
        var s = st(x);
        return '<div class="crow ' + (s.rep ? 'done' : '') + '" data-i="' + i + '" tabindex="0" role="button"><div><b>' + esc(x.complain || x.pr) + '</b>' + (x.complain && x.pr ? ' <span class="muted">PR ' + esc(x.pr) + '</span>' : '') +
          '<div>' + esc(x.place || '-') + '</div><div class="muted">' + esc([x.substation, x.feeder].filter(Boolean).join(' • ')) + (x.capacity ? ' • ' + esc(x.capacity) + ' kVA' : '') + '</div></div><div>' + chip(s) + '</div></div>';
      }).join('');
    }).catch(function (e) { if (my !== runId) return; $('cs-spin').classList.remove('on'); $('cs-res').innerHTML = ''; meta(e.message); });
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

  var q = $('cs-q'); if (!q) return;
  q.addEventListener('input', function () { clearTimeout(timer); timer = setTimeout(search, 350); });
  q.addEventListener('keydown', function (e) { if (e.key === 'Enter') { clearTimeout(timer); search(); } });
  $('cs-discom').addEventListener('change', function () { lsSet('tt_discom', this.value); search(); });
  $('cs-res').addEventListener('click', function (e) { var r = e.target.closest('.crow'); if (r) open(+r.dataset.i); });
  $('cs-res').addEventListener('keydown', function (e) { var r = e.target.closest('.crow'); if (r && e.key === 'Enter') open(+r.dataset.i); });
  $('cs-ov').addEventListener('click', function (e) { if (e.target === $('cs-ov')) close(); });
  setTimeout(loadDiscoms, 0);
})();
