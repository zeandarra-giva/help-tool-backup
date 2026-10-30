(function () {
  'use strict';

  var D = JSON.parse(document.getElementById('snapshot-data').textContent);
  var PAGE = 300;        // rows rendered per "Show more" step
  var PER_TABLE = 25;    // rows per table in "Search all tables" results
  var THEME_KEY = 'helpplus-snapshot-theme';

  function $(id) { return document.getElementById(id); }
  function esc(s) {
    return String(s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function reEsc(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

  var state = { table: null, q: '', exact: false, all: false, limit: PAGE };
  var drawerRow = null;

  // Table id -> its category, and every table id in navigation order
  var catOf = {};
  var order = [];
  D.categories.forEach(function (c) {
    c.tabs.forEach(function (t) {
      if (!catOf[t.id]) { catOf[t.id] = c; order.push(t.id); }
    });
  });
  // Start on the first top tab (like the original app), not the sidebar
  var home = (function () {
    for (var i = 0; i < D.categories.length; i++) if (!D.categories[i].sidebar) return D.categories[i].tabs[0].id;
    return order[0];
  })();

  function labelOf(id) {
    var c = catOf[id];
    for (var i = 0; c && i < c.tabs.length; i++) if (c.tabs[i].id === id) return c.tabs[i].label;
    return D.tables[id].name;
  }

  // ---- search ---------------------------------------------------------

  var hay = {};
  function rowText(id) {
    if (!hay[id]) hay[id] = D.tables[id].rows.map(function (r) { return r.join('\u0001').toLowerCase(); });
    return hay[id];
  }

  function matcher() {
    var q = state.q.trim().toLowerCase();
    if (!q) return null;
    if (state.exact) return { exact: true, q: q };
    var terms = q.split(/\s+/).filter(Boolean);
    var sorted = terms.slice().sort(function (a, b) { return b.length - a.length; });
    return { exact: false, terms: terms, re: new RegExp('(' + sorted.map(reEsc).join('|') + ')', 'gi') };
  }

  // Every word must appear somewhere in the row; exact = a whole cell equals the query
  function findRows(id, m) {
    var rows = D.tables[id].rows, out = [], i;
    if (!m) { for (i = 0; i < rows.length; i++) out.push(i); return out; }
    if (m.exact) {
      for (i = 0; i < rows.length; i++) {
        if (rows[i].some(function (v) { return v.trim().toLowerCase() === m.q; })) out.push(i);
      }
      return out;
    }
    var txt = rowText(id);
    for (i = 0; i < txt.length; i++) {
      var t = txt[i];
      if (m.terms.every(function (w) { return t.indexOf(w) !== -1; })) out.push(i);
    }
    return out;
  }

  function hl(text, m) {
    if (!m || text === '') return esc(text);
    if (m.exact) return text.trim().toLowerCase() === m.q ? '<mark>' + esc(text) + '</mark>' : esc(text);
    return text.split(m.re).map(function (part, i) {
      return i % 2 ? '<mark>' + esc(part) + '</mark>' : esc(part);
    }).join('');
  }

  // ---- rendering ------------------------------------------------------

  function tableHtml(id, idxs, m, limit) {
    var t = D.tables[id];
    var h = '<div class="table-wrap"><table><thead><tr>';
    t.cols.forEach(function (c, j) { h += '<th' + (t.num[j] ? ' class="num"' : '') + '>' + esc(c) + '</th>'; });
    h += '</tr></thead><tbody>';
    var n = Math.min(idxs.length, limit);
    for (var k = 0; k < n; k++) {
      var i = idxs[k], r = t.rows[i];
      h += '<tr data-t="' + id + '" data-r="' + i + '" tabindex="0">';
      for (var j = 0; j < r.length; j++) h += '<td' + (t.num[j] ? ' class="num"' : '') + '>' + hl(r[j], m) + '</td>';
      h += '</tr>';
    }
    return h + '</tbody></table></div>';
  }

  function renderNav() {
    var active = catOf[state.table];
    $('cats').innerHTML = D.categories.map(function (c, i) {
      if (c.sidebar) return '';
      return '<button type="button" class="cat' + (c === active ? ' active' : '') + '" data-cat="' + i + '">' + esc(c.name) + '</button>';
    }).join('');

    var sideCats = D.categories.filter(function (c) { return c.sidebar; });
    var side = sideCats.map(function (c) {
      return (sideCats.length > 1 ? '<div class="side-title">' + esc(c.name) + '</div>' : '') +
        c.tabs.map(function (t) {
          return '<button type="button" class="side-btn' + (t.id === state.table ? ' active' : '') + '" data-tab="' + t.id + '">' + esc(t.label) + '</button>';
        }).join('');
    }).join('');
    if (D.unavailable.length) {
      side += '<p class="side-note" title="' + esc(D.unavailable.join('\n')) + '">' + D.unavailable.length +
        ' quer' + (D.unavailable.length === 1 ? 'y needs' : 'ies need') + ' Access and ' + (D.unavailable.length === 1 ? 'is' : 'are') + ' not included.</p>';
    }
    $('side').innerHTML = side;
    $('side').hidden = !side;
    document.body.classList.toggle('no-side', !side);

    $('cat-title').textContent = active ? active.name : '';
    $('tabs').innerHTML = active && !active.sidebar && active.tabs.length > 1 ? active.tabs.map(function (t) {
      return '<button type="button" class="tab' + (t.id === state.table ? ' active' : '') + '" data-tab="' + t.id + '">' + esc(t.label) + '</button>';
    }).join('') : '';
  }

  function render() {
    var m = matcher();
    var res = $('results');
    $('clear').hidden = !state.q;

    if (!order.length) {
      $('table-title').textContent = '';
      $('count').textContent = '';
      res.innerHTML = '<p class="empty">No tables could be read from this database.</p>';
      return;
    }
    if (state.all && m) { renderAll(m); return; }

    var t = D.tables[state.table];
    var idxs = findRows(state.table, m);
    $('table-title').textContent = labelOf(state.table);
    $('count').textContent = m ? idxs.length + ' of ' + t.rows.length + ' rows' : t.rows.length + ' rows';

    if (!t.rows.length) {
      res.innerHTML = '<p class="empty">This table has no records.</p>';
    } else if (!idxs.length) {
      res.innerHTML = '<p class="empty">No matches for “' + esc(state.q.trim()) + '” in this table. ' +
        '<button type="button" class="link" data-act="all">Search all tables</button></p>';
    } else {
      res.innerHTML = tableHtml(state.table, idxs, m, state.limit) +
        (idxs.length > state.limit ? '<button type="button" class="more" data-act="more">Show more (' + (idxs.length - state.limit) + ' more rows)</button>' : '');
    }
  }

  function renderAll(m) {
    var total = 0, groups = 0, h = '';
    order.forEach(function (id) {
      var idxs = findRows(id, m);
      if (!idxs.length) return;
      total += idxs.length;
      groups++;
      h += '<section class="group"><div class="group-head"><h3>' + esc(labelOf(id)) + '</h3>' +
        '<span class="muted">' + idxs.length + (idxs.length === 1 ? ' match' : ' matches') + '</span>' +
        '<button type="button" class="link" data-open="' + id + '">Open table →</button></div>' +
        tableHtml(id, idxs, m, PER_TABLE) +
        (idxs.length > PER_TABLE ? '<p class="muted small">Showing the first ' + PER_TABLE + '. Open the table to see all.</p>' : '') +
        '</section>';
    });
    $('table-title').textContent = 'All tables';
    $('count').textContent = total + (total === 1 ? ' match' : ' matches') + ' in ' + groups + (groups === 1 ? ' table' : ' tables');
    $('results').innerHTML = h || '<p class="empty">No matches for “' + esc(state.q.trim()) + '” in any table.</p>';
  }

  function setTable(id) {
    if (!D.tables[id]) id = home;
    state.table = id;
    state.limit = PAGE;
    renderNav();
    render();
    $('results').scrollTop = 0;
  }

  function go(id) {
    if (state.all) { state.all = false; $('all').checked = false; }
    if (location.hash === '#' + id) setTable(id);
    else location.hash = id;
  }

  // ---- record drawer ----------------------------------------------------

  function openDrawer(id, i) {
    var t = D.tables[id], r = t.rows[i];
    drawerRow = { id: id, i: i };
    $('drawer-title').textContent = labelOf(id);
    $('drawer-body').innerHTML = t.cols.map(function (c, j) {
      var v = r[j];
      return '<div class="field"><div class="field-label">' + esc(c) + '</div>' +
        '<div class="field-value">' + (v === '' ? '<span class="muted">—</span>' : esc(v)) + '</div>' +
        (v === '' ? '' : '<button type="button" class="copy" data-copy="' + j + '">Copy</button>') + '</div>';
    }).join('');
    $('drawer').hidden = false;
    $('scrim').hidden = false;
    $('drawer-close').focus();
  }

  function closeDrawer() {
    $('drawer').hidden = true;
    $('scrim').hidden = true;
    drawerRow = null;
  }

  // ---- clipboard & toast ------------------------------------------------

  var toastTimer;
  function toast(msg) {
    var el = $('toast');
    el.textContent = msg;
    el.classList.add('show');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { el.classList.remove('show'); }, 1400);
  }

  function copyFallback(text) {
    var ta = document.createElement('textarea');
    ta.value = text;
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand('copy'); toast('Copied'); } catch (e) { toast('Copy is blocked in this browser'); }
    ta.remove();
  }

  function copy(text) {
    if (navigator.clipboard && window.isSecureContext) {
      navigator.clipboard.writeText(text).then(function () { toast('Copied'); }, function () { copyFallback(text); });
    } else {
      copyFallback(text);
    }
  }

  // ---- theme --------------------------------------------------------------

  var root = document.documentElement;
  function applyTheme(t) {
    root.setAttribute('data-theme', t);
    $('theme').textContent = t === 'dark' ? '☀ Light mode' : '☾ Dark mode';
  }
  try { var saved = localStorage.getItem(THEME_KEY); if (saved) root.setAttribute('data-theme', saved); } catch (e) { }
  applyTheme(root.getAttribute('data-theme') === 'light' ? 'light' : 'dark');

  // ---- events -------------------------------------------------------------

  document.addEventListener('click', function (ev) {
    var el = ev.target.closest('[data-cat],[data-tab],[data-open],[data-act],[data-copy],tr[data-t]');
    if (!el) return;
    if (el.hasAttribute('data-cat')) go(D.categories[+el.getAttribute('data-cat')].tabs[0].id);
    else if (el.hasAttribute('data-tab')) go(el.getAttribute('data-tab'));
    else if (el.hasAttribute('data-open')) go(el.getAttribute('data-open'));
    else if (el.hasAttribute('data-copy')) copy(D.tables[drawerRow.id].rows[drawerRow.i][+el.getAttribute('data-copy')]);
    else if (el.hasAttribute('data-act')) {
      var act = el.getAttribute('data-act');
      if (act === 'more') { state.limit += PAGE; render(); }
      if (act === 'all') { state.all = true; $('all').checked = true; render(); }
    } else openDrawer(el.getAttribute('data-t'), +el.getAttribute('data-r'));
  });

  var debounce;
  $('q').addEventListener('input', function () {
    clearTimeout(debounce);
    debounce = setTimeout(function () { state.q = $('q').value; state.limit = PAGE; render(); }, 120);
  });
  $('clear').addEventListener('click', function () {
    $('q').value = '';
    state.q = '';
    render();
    $('q').focus();
  });
  $('exact').addEventListener('change', function () { state.exact = this.checked; render(); });
  $('all').addEventListener('change', function () { state.all = this.checked; render(); });
  $('scrim').addEventListener('click', closeDrawer);
  $('drawer-close').addEventListener('click', closeDrawer);
  $('copy-all').addEventListener('click', function () {
    if (!drawerRow) return;
    var t = D.tables[drawerRow.id], r = t.rows[drawerRow.i];
    copy(t.cols.map(function (c, j) { return c + ': ' + r[j]; }).join('\n'));
  });
  $('theme').addEventListener('click', function () {
    var next = root.getAttribute('data-theme') === 'dark' ? 'light' : 'dark';
    applyTheme(next);
    try { localStorage.setItem(THEME_KEY, next); } catch (e) { }
  });
  $('logo').addEventListener('click', function (ev) { ev.preventDefault(); go(home); });

  document.addEventListener('keydown', function (ev) {
    var typing = /^(INPUT|TEXTAREA|SELECT)$/.test(document.activeElement.tagName);
    if (ev.key === '/' && !typing) { ev.preventDefault(); $('q').focus(); $('q').select(); }
    else if (ev.key === 'Escape') {
      if (!$('drawer').hidden) closeDrawer();
      else if (document.activeElement === $('q') && $('q').value) $('clear').click();
    } else if (ev.key === 'Enter' && document.activeElement.matches('tr[data-t]')) {
      var tr = document.activeElement;
      openDrawer(tr.getAttribute('data-t'), +tr.getAttribute('data-r'));
    }
  });

  window.addEventListener('hashchange', function () { setTable(location.hash.slice(1)); });

  // ---- start ----------------------------------------------------------------

  $('logo').textContent = D.title;
  $('stamp').textContent = 'Snapshot · ' + D.stamp;
  document.title = D.title + ' – snapshot ' + D.stamp;
  setTable(location.hash.slice(1));
})();
