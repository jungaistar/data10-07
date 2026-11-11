/* Quality Planner — 화면 (대시보드 · 불량 등록 · 파일 업로드(선택) · 불량 분석 · MASTER PLAN · 추진현황 · 보고서 출력 · 설정)
   계산은 js/logic.js, 저장은 js/store.js, AI 연결은 js/ai-endpoint.js · js/ai-panel.js 에 있습니다. */
(function () {
  'use strict';
  var L = window.SQLogic;
  /* 판 맞추기 — 브라우저가 옛 logic.js 와 새 app.js 를 섞어 쓰면(Pages 캐시) 버튼과 결과가 어긋날 수 있어, 섞였으면 계산하지 않고 새로고침을 안내합니다.
     판은 scripts/bump-version.mjs 로 index.html(?v=) · logic.js 와 함께 올립니다. */
  var APP_VERSION = '2026-11-11.1';
  if (!L || L.VERSION !== APP_VERSION) {
    var m0 = document.getElementById('main') || document.body;
    m0.innerHTML = '<section class="card" id="versionMismatch"><h1>새 판이 나왔습니다 — 새로고침(Ctrl+F5)해 주세요</h1><p class="alert warn">이 브라우저가 옛 파일과 새 파일을 섞어 불러왔습니다(js/logic.js ' +
      (L && L.VERSION ? L.VERSION : '옛 판') + ' · 화면 ' + APP_VERSION + '). 이대로 계산하면 버튼과 결과가 어긋날 수 있어 멈췄습니다.</p>' +
      '<p>Windows 는 <b>Ctrl+F5</b>, Mac 은 <b>Cmd+Shift+R</b> 을 한 번 눌러 주세요. 쌓아 둔 자료는 지워지지 않습니다.</p></section>';
    return;
  }
  (function () {
    var ft = document.querySelector('.site-footer .inner') || document.querySelector('.site-footer'); if (!ft) return;
    var s = document.createElement('span'); s.className = 'ver'; s.textContent = ' · 판 ' + APP_VERSION + (window.__SINGLE_FILE__ ? ' (HTML 한 파일)' : ''); ft.appendChild(s);
    if (!window.__SINGLE_FILE__) { ft.appendChild(document.createTextNode(' · ')); var a = document.createElement('a'); a.href = '#/settings'; a.textContent = 'HTML 파일로 받기'; a.addEventListener('click', function () { setTimeout(function () { var c = document.getElementById('singleCard'); if (c) c.scrollIntoView({ block: 'start' }); }, 50); }); ft.appendChild(a); }
  })();
  var S = window.SQStore;
  var Sample = window.SQSample;
  var XLSX = window.XLSX;
  var XO = window.SQXlsx;
  var PAGE = 50;
  var SINGLE = !!window.__SINGLE_FILE__;   // scripts/build-single.mjs 가 만든 「HTML 한 파일」 판이면 true
  var SINGLE_NAME = 'data10-07_협력사공정불량_개선관리.html';

  var db = S.loadDb();
  var ui = {
    an: { supplier: '', campus: '', from: '', to: '', type: '', q: '', onlyUnc: false, page: 0 },
    plan: { supplier: '', scale: 'week', aiText: '', parsed: null, mode: 'replace', showPrompt: false, compact: false },
    prog: { supplier: '', phase: '', scale: 'month' },
    reply: { file: null, parsed: null, diff: null, take: {}, error: '' },
    rep: { supplier: '', campus: '' },
    mail: { due: '', per: '' },
    entry: { editId: '', keep: { date: '', supplier: '', campus: '' }, paste: '', pasted: null, supplier: '', campus: '', type: '', q: '', page: 0 },
    imp: null,
    impMode: 'add'
  };

  // ── 작은 도구 ─────────────────────────────────────────────
  function h(tag, attrs) {
    var el = document.createElement(tag);
    if (attrs) Object.keys(attrs).forEach(function (k) {
      var v = attrs[k];
      if (v == null || v === false) return;
      if (k === 'class') el.className = v;
      else if (k === 'text') el.textContent = v;
      else if (k === 'value') el.value = v;
      else if (k.slice(0, 2) === 'on' && typeof v === 'function') el.addEventListener(k.slice(2), v);
      else if (v === true) el.setAttribute(k, '');
      else el.setAttribute(k, v);
    });
    for (var i = 2; i < arguments.length; i++) add(el, arguments[i]);
    return el;
  }
  function add(el, c) {
    if (c == null || c === false) return;
    if (Array.isArray(c)) { c.forEach(function (x) { add(el, x); }); return; }
    el.appendChild(typeof c === 'string' || typeof c === 'number' ? document.createTextNode(String(c)) : c);
  }
  function fmt(n) { return n == null || n === '' ? '' : Number(n).toLocaleString('ko-KR', { maximumFractionDigits: 1 }); }
  function toast(msg, isError) {
    var t = document.getElementById('toast');
    t.textContent = msg;
    t.className = 'toast' + (isError ? ' error' : '');
    t.hidden = false;
    clearTimeout(toast._t);
    toast._t = setTimeout(function () { t.hidden = true; }, 3800);
  }
  function save() {
    if (!S.saveDb(db)) {
      document.getElementById('storeBanner').hidden = false;
      if (S.lastError() === 'quota') toast('브라우저 저장 공간이 찼습니다. 설정에서 JSON 백업을 받아 두세요.', true);
    }
  }
  function today() { return L.toDateStr(new Date()); }
  function nowIso() { var d = new Date(); return today() + 'T' + ('0' + d.getHours()).slice(-2) + ':' + ('0' + d.getMinutes()).slice(-2) + ':' + ('0' + d.getSeconds()).slice(-2); }
  function nextMonday() { var t = today(), m = L.mondayOf(t); return m === t ? t : L.addDays(m, 7); }
  function field(label, input, hint, cls) {
    return h('label', { class: 'field' + (cls ? ' ' + cls : '') }, h('span', null, label), input, hint ? h('small', { class: 'hint' }, hint) : null);
  }
  function select(options, value, attrs) {
    var s = h('select', attrs || {});
    options.forEach(function (o) {
      var v = Array.isArray(o) ? o[0] : o, t = Array.isArray(o) ? o[1] : o;
      s.appendChild(h('option', { value: v, selected: String(v) === String(value) }, t));
    });
    return s;
  }
  function dateInput(value, attrs) { return h('input', Object.assign({ type: 'date', value: value || '' }, attrs || {})); }
  function datalist(id, values) { return h('datalist', { id: id }, values.map(function (v) { return h('option', { value: v }); })); }
  var onDialogClose = null;
  function openDialog(title, content, actions) {
    var d = document.getElementById('dialog');
    document.getElementById('dialogTitle').textContent = title;
    var c = document.getElementById('dialogContent'); c.textContent = ''; add(c, content);
    var a = document.getElementById('dialogActions'); a.textContent = ''; add(a, actions);
    if (!d.open) d.showModal();
  }
  function closeDialog() { var d = document.getElementById('dialog'); if (d.open) d.close(); }
  document.getElementById('dialog').addEventListener('close', function () { var f = onDialogClose; onDialogClose = null; if (f) f(); });
  function confirmBox(title, text, okLabel, onOk, danger) {
    openDialog(title, h('p', { class: 'pre' }, text), [
      h('button', { type: 'button', class: 'btn', onclick: closeDialog }, '그만두기'),
      h('button', { type: 'button', class: 'btn ' + (danger ? 'btn-danger' : 'btn-primary'), onclick: function () { closeDialog(); onOk(); } }, okLabel)
    ]);
  }
  function download(name, blob) {
    var a = h('a', { href: URL.createObjectURL(blob), download: name });
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }
  /* 엑셀 쓰기 — js/xlsx-out.js 가 머리글 굵게·칸 색·날짜 형식·열 폭·틀 고정을 붙입니다(외부 서버 없이) */
  function writeXlsx(name, sheets, meta) {
    var out = XO.build(XLSX, sheets, meta || {});
    download(name, new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  }
  function listMeta(aoa, widths) { return { head: [0], freeze: { rows: 1 }, filter: 'A1:' + XO.colName(Math.max(0, (aoa[0] || []).length - 1)) + Math.max(1, aoa.length), widths: widths }; }
  function copyText(s) {
    function fallback() {
      var ta = h('textarea', { style: 'position:fixed;left:-9999px;top:0' });
      ta.value = s; document.body.appendChild(ta); ta.select();
      var ok = false;
      try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
      ta.remove();
      toast(ok ? '복사했습니다' : '복사하지 못했습니다. 글상자를 눌러 직접 복사하세요', !ok);
    }
    if (navigator.clipboard && window.isSecureContext) navigator.clipboard.writeText(s).then(function () { toast('복사했습니다'); }, fallback);
    else fallback();
  }
  function tag() { return db._sample ? '_예시데이터' : ''; }
  function all() { return L.enrich(db.records, db.types); }
  function suppliers() { return L.suppliersOf(db.records); }
  /* 진행상태 표시(4차) — 진행중 · 완료 · 미완료 · 보류 네 가지. 완료·보류는 적은 그대로, 나머지는 오늘 날짜로(마감이 지나면 미완료) (logic.js phaseOf) */
  function statusPill(r) {
    var ph = L.phaseOf(r, today());
    return h('span', { class: 'pill st-' + ph }, L.PHASE[ph], L.notStarted(r, today()) ? h('small', { class: 'pill-sub' }, ' 시작 전') : null);
  }
  /* 캠퍼스 — 울산 · 인천 RAW DATA 를 따로 올려 합칩니다(4차). 데이터에 캠퍼스가 하나도 없으면 고르기 칸을 숨깁니다 */
  function hasCampus() { return db.records.some(function (r) { return r.campus; }); }
  function campusText(c) { var n = L.campusCounts(c || db.records); return L.CAMPUSES.map(function (k) { return L.campusLabel(k) + ' ' + n[k]; }).join(' · ') + (n[''] ? ' · 구분 없음 ' + n[''] : ''); }
  function campusSelect(value, onChange, allLabel) {
    var s = select([['', allLabel || '전체 캠퍼스']].concat(L.CAMPUSES.map(function (c) { return [c, L.campusLabel(c)]; })).concat([['_none', '구분 없음(직접 등록 등)']]), value, { 'aria-label': '캠퍼스' });
    s.addEventListener('change', function () { onChange(s.value); });
    return s;
  }
  /* 개선 필요 항목의 구분 — WORST n · 동일품번 */
  function issueBadge(i) {
    return i.kind === 'part' ? h('span', { class: 'flag part' }, '동일품번') : h('span', { class: 'flag w' + i.rank }, i.label);
  }
  function emptyNotice(main, title) {
    main.appendChild(pageHead(title));
    main.appendChild(h('section', { class: 'card' },
      h('p', null, '아직 공정불량 데이터가 없습니다. 「불량 등록」에서 한 건씩 적거나 엑셀 표를 복사해 붙여 넣으세요. 예시 데이터로 흐름을 먼저 볼 수도 있습니다.'),
      h('div', { class: 'btn-row' }, h('a', { class: 'btn btn-primary', href: '#/entry' }, '불량 등록으로'), sampleButton(false))));
  }
  function pageHead(title, sub, right) {
    return h('div', { class: 'page-head' }, h('div', { class: 'ph-text' }, h('h1', null, title), sub ? h('p', { class: 'note' }, sub) : null), right || null);
  }

  // ── 아이콘·차트 부품 ──────────────────────────────────────
  var SVGNS = 'http://www.w3.org/2000/svg';
  function sv(tagName, attrs) {
    var el = document.createElementNS(SVGNS, tagName);
    Object.keys(attrs || {}).forEach(function (k) { el.setAttribute(k, attrs[k]); });
    for (var i = 2; i < arguments.length; i++) { var c = arguments[i]; if (c == null) continue; el.appendChild(typeof c === 'string' ? document.createTextNode(c) : c); }
    return el;
  }
  var ICON = {
    doc: 'M6 3h9l4 4v14H6zM9 12h7M9 16h5M14 3v5h5',
    factory: 'M3 20V10l5 3V10l5 3V6h3l1 4h3v10zM7 17h2M12 17h2M17 17h1',
    alert: 'M12 4 2.8 19.5h18.4zM12 10v4.5M12 17.2v.3',
    check: 'M4 4h16v16H4zM8 12.5l3 3 5-6',
    chev: 'm9 6 6 6-6 6',
    tick: 'm5 12.5 4.5 4.5L19 7.5',
    chip: 'M7 7h10v10H7zM10 10h4v4h-4zM9.5 3v4M14.5 3v4M9.5 17v4M14.5 17v4M3 9.5h4M3 14.5h4M17 9.5h4M17 14.5h4',
    layers: 'M12 3 3 8l9 5 9-5zM3 13l9 5 9-5M3 17.5l9 5 9-5',
    list: 'M4 6h16M4 12h16M4 18h10',
    cal: 'M4 6h16v14H4zM4 10h16M8 3v5M16 3v5',
    clock: 'M12 3.5a8.5 8.5 0 1 0 0 17a8.5 8.5 0 1 0 0-17M12 7.5V12l3 2',
    pause: 'M12 3.5a8.5 8.5 0 1 0 0 17a8.5 8.5 0 1 0 0-17M10 9v6M14 9v6',
    plus: 'M12 5v14M5 12h14',
    down: 'M12 4v11M7.5 10.5 12 15l4.5-4.5M4 19h16',
    redo: 'M4 12a8 8 0 0 1 14-5.3L20 9M20 4v5h-5M20 12a8 8 0 0 1-14 5.3L4 15M4 20v-5h5'
  };
  function icon(name, cls) { return sv('svg', { viewBox: '0 0 24 24', class: 'ico' + (cls ? ' ' + cls : ''), 'aria-hidden': 'true' }, sv('path', { d: ICON[name] })); }
  function kpi(title, iconName, value, unit, note, tone) {
    return h('div', { class: 'dk' + (tone ? ' tone-' + tone : '') },
      h('span', { class: 'dk-ico' }, icon(iconName)),
      h('div', { class: 'dk-body' }, h('div', { class: 'dk-t' }, title), h('div', { class: 'dk-v' }, value, unit ? h('small', null, unit) : null), note ? h('div', { class: 'dk-note' }, note) : null));
  }
  function panel(title, more, body, cls) {
    return h('section', { class: 'dp' + (cls ? ' ' + cls : '') },
      h('div', { class: 'dp-head' }, h('h2', null, title), more ? h('a', { class: 'dp-more', href: more[0] }, more[1] || '더보기', icon('chev')) : null),
      body);
  }
  // 가로 막대 — 글자는 HTML 이라 좁은 화면에서도 작아지지 않습니다
  function hbars(items, opt) {
    var max = Math.max.apply(null, items.map(function (x) { return x.value; }).concat([1]));
    return h('ul', { class: 'hbars' + (opt && opt.cls ? ' ' + opt.cls : ''), 'aria-label': opt && opt.label },
      items.map(function (x) {
        var name = x.href ? h('a', { href: x.href, class: 'hb-k' }, x.key) : h('span', { class: 'hb-k' }, x.key);
        return h('li', { class: 'hb' + (x.cls ? ' ' + x.cls : '') }, name,
          h('span', { class: 'hb-track' }, h('span', { class: 'hb-fill', style: 'width:' + (x.value ? Math.max(2, x.value / max * 100) : 0) + '%' })),
          h('span', { class: 'hb-v' }, fmt(x.value) + (x.unit || '건'), x.sub ? h('small', null, ' ' + x.sub) : null));
      }));
  }
  function vbars(items, opt) {
    var max = Math.max.apply(null, items.map(function (x) { return x.value; }).concat([1]));
    return h('div', { class: 'vbars', role: 'img', 'aria-label': opt.label },
      items.map(function (x) {
        return h('div', { class: 'vbar', title: x.tip || null },
          h('span', { class: 'vbar-v' }, fmt(x.value)),
          h('span', { class: 'vbar-track' }, h('span', { class: 'vbar-fill', style: 'height:' + (x.value ? Math.max(3, x.value / max * 100) : 0) + '%' })),
          h('span', { class: 'vbar-k' }, x.key));
      }));
  }
  function monthBars(months, label) {
    var list = months.slice(-12);
    return vbars(list.map(function (m) { return { key: +m.month.slice(5) + '월', value: m.count, tip: m.month + ' · ' + m.count + '건' }; }), { label: label });
  }
  function paretoBars(byType) {
    return hbars(byType.map(function (t) { return { key: t.type, value: t.count, sub: '(' + t.pct + '% · 누적 ' + t.cum + '%)', cls: t.type === L.UNCLASSIFIED ? 'unc' : '' }; }), { label: '불량유형별 건수 (파레토)', cls: 'wide' });
  }
  function stackBar(st) {
    var tot = st.total || 1;
    return h('span', { class: 'stack', role: 'img', 'aria-label': L.PHASE_ORDER.map(function (k) { return L.PHASE[k] + ' ' + st[k]; }).join(', ') },
      L.PHASE_ORDER.map(function (k) { return st[k] ? h('span', { class: 'stack-' + k, style: 'width:' + (st[k] / tot * 100) + '%' }) : null; }));
  }
  /* 개선과제 진행현황 — 전체 · 진행중 · 완료 · 미완료 · 보류 (수강생 그림의 「개선 MASTER PLAN 현황」) */
  function phaseTiles(st, href) {
    function tile(k, label, n, cls) { return h(href ? 'a' : 'div', { class: 'pt ' + cls, href: href ? href + (k ? '/' + k : '') : null }, h('span', { class: 'pt-k' }, label), h('span', { class: 'pt-v' }, String(n))); }
    return h('div', { class: 'ptiles' },
      tile('', '전체 과제', st.total, 'pt-all'), tile('doing', '진행중', st.doing, 'pt-doing'), tile('done', '완료', st.done, 'pt-done'), tile('late', '미완료', st.late, 'pt-late'), tile('hold', '보류', st.hold, 'pt-hold'),
      h('p', { class: 'note pt-note' }, '완료율 ' + st.rate + '% — 진행중 = 개선 마감 전(그 가운데 시작 전 ' + st.notStarted + '건), 미완료 = 마감이 지났는데 완료가 아님, 완료·보류 = 적은 그대로'));
  }

  // ── 라우팅 ────────────────────────────────────────────────
  function routeParts() { return location.hash.replace(/^#\/?/, '').split('/'); }
  function route() { return routeParts()[0] || 'home'; }
  function routeArg() { var p = routeParts()[1]; try { return p ? decodeURIComponent(p) : ''; } catch (e) { return ''; } }
  function render() {
    var r = route();
    var main = document.getElementById('main');
    main.textContent = '';
    document.querySelectorAll('.sidebar a[data-route]').forEach(function (a) {
      if (a.getAttribute('data-route') === r) a.setAttribute('aria-current', 'page'); else a.removeAttribute('aria-current');
    });
    document.getElementById('sampleBanner').hidden = !db._sample;
    if (!S.available()) document.getElementById('storeBanner').hidden = false;
    document.body.setAttribute('data-route', r);
    if (r === 'entry') renderEntry(main);
    else if (r === 'upload') renderUpload(main);
    else if (r === 'analysis') renderAnalysis(main);
    else if (r === 'plan') renderPlan(main);
    else if (r === 'progress') renderProgress(main);
    else if (r === 'report') renderReport(main);
    else if (r === 'reply') renderReply(main);
    else if (r === 'settings') renderSettings(main);
    else renderHome(main);
  }
  window.addEventListener('hashchange', function () { setMenu(false); closeDrawer(true); render(); window.scrollTo(0, 0); });

  // 좁은 화면(1024px 미만)에서는 왼쪽 메뉴를 「메뉴」 단추로 여닫습니다.
  var menuBtn = document.getElementById('menuBtn'), sidebar = document.getElementById('sidebar'), backdrop = document.getElementById('sideBackdrop');
  function setMenu(open) {
    document.body.classList.toggle('menu-open', open);
    menuBtn.setAttribute('aria-expanded', open ? 'true' : 'false');
    menuBtn.setAttribute('aria-label', open ? '메뉴 닫기' : '메뉴 열기');
    backdrop.hidden = !open;
    if (open) { var first = sidebar.querySelector('a'); if (first) first.focus(); }
  }
  menuBtn.addEventListener('click', function () { setMenu(!document.body.classList.contains('menu-open')); });
  backdrop.addEventListener('click', function () { setMenu(false); menuBtn.focus(); });
  document.addEventListener('keydown', function (e) {
    if (e.key !== 'Escape') return;
    if (!document.getElementById('drawer').hidden) { closeDrawer(); return; }
    if (document.body.classList.contains('menu-open')) { setMenu(false); menuBtn.focus(); }
  });
  sidebar.addEventListener('click', function (e) { if (e.target.closest('a')) setMenu(false); });
  var WEEK = ['일', '월', '화', '수', '목', '금', '토'];
  function nowText() {
    var d = new Date();
    function p(n) { return (n < 10 ? '0' : '') + n; }
    document.getElementById('nowText').textContent = d.getFullYear() + '. ' + p(d.getMonth() + 1) + '. ' + p(d.getDate()) + ' (' + WEEK[d.getDay()] + ') ' + p(d.getHours()) + ':' + p(d.getMinutes());
  }
  nowText(); setInterval(nowText, 30000);
  // 간트는 700px 을 사이에 두고 칸 구성이 바뀌므로 그 경계를 넘을 때만 다시 그립니다
  var wasNarrow = document.documentElement.clientWidth < 700;
  window.addEventListener('resize', function () {
    clearTimeout(render._r);
    render._r = setTimeout(function () {
      var n = document.documentElement.clientWidth < 700;
      if (n !== wasNarrow) { wasNarrow = n; if (route() === 'plan' || route() === 'progress') render(); }
    }, 200);
  });

  // ── 예시 데이터 ───────────────────────────────────────────
  function sampleButton(primary) {
    return h('button', { type: 'button', class: 'btn' + (primary ? ' btn-primary' : ''), onclick: loadSample }, '예시 데이터로 시작');
  }
  function loadSample() {
    function go() {
      var keep = { types: db.types, depts: db.depts, rules: db.rules, mask: db.mask, mapping: db.mapping, mail: db.mail };
      db = L.emptyDb();
      Object.assign(db, keep);
      // 4차: 예시도 울산 · 인천캠퍼스 RAW DATA(열 구성이 다른 두 시스템 모양)를 각각 읽어 합칩니다
      var recs = [], seq = 0, files = [];
      L.CAMPUSES.forEach(function (c) {
        var m = L.replaceCampus(recs, c, L.parseSheet(Sample.campusAoa(c)).records, seq);
        recs = m.records; seq = m.seq;
        files.push({ name: '예시 데이터(가상) · ' + L.campusLabel(c), campus: c, at: nowIso(), added: m.added, dup: m.dup });
      });
      db.records = recs; db.seq = seq; db._sample = true;
      db.aliases = L.updateAliases(db.aliases, db.records);
      db.files = files;
      save();
      toast('예시 데이터 ' + recs.length + '건을 넣었습니다(협력사 A~F, 가상 · ' + campusText() + ').');
      if (route() !== 'analysis') location.hash = '#/analysis'; else render();
    }
    if (db.records.length && !db._sample) confirmBox('예시 데이터로 바꾸기', '지금 있는 Raw data ' + db.records.length + '건과 Master Plan 을 지우고 예시 데이터를 넣습니다.\n필요하면 먼저 설정에서 JSON 백업을 받아 두세요.', '바꾸기', go, true);
    else go();
  }

  // ══ 대시보드 ═════════════════════════════════════════════
  function renderHome(main) {
    var rows = all();
    var b = L.dateBounds(rows);
    main.appendChild(h('section', { class: 'hero' },
      h('div', { class: 'hero-text' },
        h('h1', null, '협력사 품질 현황'),
        h('p', null, '협력사 공정불량을 유형별로 나누어 보고, 개선이 필요한 항목부터 Master Plan 으로 관리합니다.')),
      b.min ? h('span', { class: 'hero-chip' }, icon('cal'), b.min + ' ~ ' + b.max) : null));
    if (!rows.length) {
      main.appendChild(h('section', { class: 'dp dash-empty' },
        h('h2', null, '아직 공정불량 데이터가 없습니다'),
        h('p', null, '「불량 등록」에서 협력사·품번·발생일·불량유형·내용·수량·조치를 바로 적거나, 엑셀에서 표를 복사해 붙여 넣으면 이 화면의 지표·차트가 채워집니다. 먼저 예시 데이터(가상)로 흐름을 볼 수도 있습니다.'),
        h('div', { class: 'btn-row' }, h('a', { class: 'btn btn-primary', href: '#/entry' }, '불량 등록'), sampleButton(false))));
      main.appendChild(flowGuide());
      return;
    }
    var an = L.analyze(rows, {}, db.rules);
    var allPlan = [];
    Object.keys(db.plans).forEach(function (s) { allPlan = allPlan.concat(db.plans[s].rows); });
    var ps = L.planStats(allPlan, today());
    var last3 = L.analyze(rows, { from: L.monthAdd(L.monthOf(an.to), -2) + '-01', to: an.to }, db.rules);
    main.appendChild(h('section', { class: 'dkpis' },
      kpi('불량 발생 건수', 'doc', fmt(an.total), '건', hasCampus() ? campusText() : '최근 3개월 ' + last3.total + '건'),
      kpi('협력사', 'factory', an.suppliers.length, '곳', an.unclassified ? '유형 미분류 ' + an.unclassified + '건' : '모든 건 유형 분류됨'),
      kpi('개선 필요 항목', 'alert', an.issues.length, '건', 'WORST ' + an.worst.length + ' · 동일품번 ' + an.repeatParts.length, 'warn'),
      kpi('개선과제', 'check', ps.total, '건', ps.total ? '진행중 ' + ps.doing + ' · 완료 ' + ps.done + ' · 미완료 ' + ps.late + ' · 보류 ' + ps.hold : 'Master Plan 을 아직 만들지 않았습니다', 'ok')));
    if (ps.total) main.appendChild(panel('개선 MASTER PLAN 현황', ['#/progress', '추진현황'], phaseTiles(ps, '#/progress'), 'p-tiles'));

    main.appendChild(h('div', { class: 'dcharts' },
      panel('협력사별 불량 건수', ['#/analysis', '불량 분석'], hbars(an.suppliers.map(function (s) { return { key: s.name, value: s.count, href: '#/analysis/' + encodeURIComponent(s.name) }; }), { label: '협력사별 불량 건수' })),
      panel('불량유형 파레토', ['#/analysis', '불량 분석'], paretoBars(an.byType)),
      panel('월별 불량 추이', null, h('div', null, h('p', { class: 'dp-sub' }, '최근 12개월 · 막대 = 건수'), monthBars(an.months, '월별 불량 건수')))));

    var tb = h('tbody');
    an.issues.slice(0, 8).forEach(function (i) {
      tb.appendChild(h('tr', null,
        h('td', { class: 'nw' }, i.supplier), h('td', { class: 'nw' }, issueBadge(i)), h('td', { class: 'nw' }, L.issueTitle(i)), h('td', { class: 'num' }, i.count + '건'),
        h('td', null, h('a', { class: 'btn btn-sm2', href: '#/plan/' + encodeURIComponent(i.supplier) }, '계획'))));
    });
    var issues = panel('개선 필요 항목 TOP 8', ['#/analysis', '전체 보기'], an.issues.length
      ? h('div', { class: 'table-wrap flat' }, h('table', { class: 'list dash-table' }, h('thead', null, h('tr', null, ['협력사', '구분', '불량유형 · 품번', '건수', 'Master Plan'].map(function (t) { return h('th', null, t); }))), tb))
      : h('p', { class: 'note' }, '아직 개선 필요 항목이 없습니다.'), 'a-issues');

    var pt = h('tbody');
    suppliers().forEach(function (s) {
      var p = db.plans[s]; if (!p || !p.rows.length) return;
      var st = L.planStats(p.rows, today());
      pt.appendChild(h('tr', { class: 'clickable', tabindex: '0', onclick: function () { location.hash = '#/plan/' + encodeURIComponent(s); }, onkeydown: function (e) { if (e.key === 'Enter') location.hash = '#/plan/' + encodeURIComponent(s); } },
        h('td', { class: 'nw' }, s), h('td', { class: 'stack-cell' }, stackBar(st)), h('td', { class: 'num' }, st.total), h('td', { class: 'num' }, st.doing), h('td', { class: 'num' }, st.done), h('td', { class: 'num' }, st.late ? h('span', { class: 'late' }, String(st.late)) : '0'), h('td', { class: 'num' }, st.hold)));
    });
    var prog = panel('협력사별 추진현황', ['#/progress', '추진현황'], pt.children.length
      ? h('div', null, h('div', { class: 'table-wrap flat' }, h('table', { class: 'list dash-table' }, h('thead', null, h('tr', null, ['협력사', '진행', '전체', '진행중', '완료', '미완료', '보류'].map(function (t) { return h('th', null, t); }))), pt)), legend())
      : h('p', { class: 'note' }, '아직 Master Plan 이 없습니다. 「MASTER PLAN」에서 협력사를 골라 「분석 및 초안 생성」을 누르세요.'), 'a-prog');
    main.appendChild(h('div', { class: 'dash-low' }, issues, prog));
    main.appendChild(flowGuide());
  }
  function legend() {
    return h('ul', { class: 'legend' }, L.PHASE_ORDER.map(function (k) { return h('li', null, h('i', { class: 'lg-' + k }), L.PHASE[k]); }));
  }
  function flowGuide() {
    var steps = [
      ['#/entry', '불량 등록', '화면에서 바로 적거나 엑셀 표를 복사해 붙여 넣습니다'],
      ['#/analysis', '불량 분석', '불량유형별 WORST 1·2·3 과 동일품번 2건 이상을 개선 필요 항목으로 찾습니다'],
      ['#/plan', 'MASTER PLAN', '유형별 기본 대책으로 초안 → 직접 다듬기(Copilot 등) → 간트로 계획 확정'],
      ['#/upload', '캠퍼스 RAW DATA', '울산 · 인천캠퍼스 RAW DATA 를 각각 올려 합칩니다(선택 — 직접 등록만 해도 됨)'],
      ['#/report', 'Excel 추출 · 메일', '「불량목록」·「MASTER PLAN」 두 시트를 만들고, 아웃룩 메일 초안(.eml)으로 협력사에 보냅니다'],
      ['#/reply', '진행사항 업로드', '협력사가 고쳐 회신한 Excel 을 올려 바뀐 칸만 반영합니다'],
      ['#/progress', '추진현황', '진행중 · 완료 · 미완료 · 보류를 협력사별로 봅니다']
    ];
    return panel('일하는 순서', null, h('ol', { class: 'flow' }, steps.map(function (s, i) {
      return h('li', null, h('a', { href: s[0] }, h('span', { class: 'flow-n' }, String(i + 1)), h('span', { class: 'flow-t' }, h('strong', null, s[1]), h('span', null, s[2])), icon('chev', 'q-chev')));
    })), 'flow-panel');
  }

  // ══ 불량 등록 (기본 경로 — 파일 없이 화면에서 바로 입력 · 엑셀 붙여넣기) ══════════
  function typeOptions(withAuto) {
    return (withAuto ? [['', '(비워 두면 내용 키워드로 자동 분류)']] : []).concat(L.typeNames(db.types).map(function (n) { return [n, n]; }));
  }
  /* 예시 데이터가 들어 있으면 먼저 비우고 새로 시작합니다(사전·규칙·AI 설정은 남김) */
  function leaveSample(then) {
    if (!db._sample) { then(); return; }
    confirmBox('예시 데이터 비우기', '지금 들어 있는 것은 예시 데이터(가상)입니다.\n예시 데이터와 그 Master Plan 을 지우고, 입력한 내용으로 새로 시작합니다.', '비우고 시작', function () {
      var keep = { types: db.types, depts: db.depts, rules: db.rules, mask: db.mask, mapping: db.mapping, mail: db.mail };
      db = L.emptyDb(); Object.assign(db, keep);
      then();
    });
  }
  function renderEntry(main) {
    var st = ui.entry;
    main.appendChild(pageHead('불량 등록', '협력사 공정불량을 이 화면에서 바로 적거나, 엑셀에서 표를 복사해 붙여 넣습니다. 입력한 내용은 이 브라우저에만 저장되고 서버로 보내지 않습니다.'));
    if (db._sample) main.appendChild(h('p', { class: 'alert info' }, '지금은 예시 데이터(가상)가 들어 있습니다. 첫 건을 등록하거나 붙여 넣으면 예시 데이터를 비우고 새로 시작합니다.'));
    main.appendChild(entryForm());
    main.appendChild(pasteCard());
    main.appendChild(entryList());
    main.appendChild(backupCard());
  }
  function entryForm() {
    var st = ui.entry;
    var orig = st.editId ? db.records.filter(function (r) { return r.id === st.editId; })[0] : null;
    if (st.editId && !orig) st.editId = '';
    var v = orig ? { date: orig.date, supplier: orig.supplier, partNo: orig.partNo, partName: orig.partName, type: orig.typeManual || '', problem: orig.problem, qty: orig.qty == null ? '' : String(orig.qty), action: orig.action || '' }
      : { date: st.keep.date || today(), supplier: st.keep.supplier || '', partNo: '', partName: '', type: '', problem: '', qty: '', action: '', campus: st.keep.campus || '' };
    if (orig) v.campus = orig.campus || '';
    var camp = select([['', '구분 없음']].concat(L.CAMPUSES.map(function (c) { return [c, L.campusLabel(c)]; })), v.campus, { 'aria-label': '캠퍼스' });
    var date = dateInput(v.date, { required: true, 'aria-required': 'true' });
    var sup = h('input', { value: v.supplier, list: 'dl-sup', placeholder: '예: 협력사 A', 'aria-required': 'true', autocomplete: 'off' });
    var pn = h('input', { value: v.partNo, placeholder: '예: TP-0001', autocomplete: 'off' });
    var pname = h('input', { value: v.partName, placeholder: '예: 브래킷', autocomplete: 'off' });
    var type = select(typeOptions(true), v.type, { 'aria-label': '불량유형' });
    var qty = h('input', { type: 'number', min: '0', step: 'any', value: v.qty, inputmode: 'decimal' });
    var prob = h('textarea', { rows: '2', placeholder: '예: 브래킷 용접부 균열 발생', 'aria-required': 'true' }); prob.value = v.problem;
    var act = h('textarea', { rows: '2', placeholder: '예: 전수 선별 후 재작업, 협력사 반품' }); act.value = v.action;
    var hint = h('small', { class: 'hint' });
    function showAuto() {
      if (type.value) { hint.textContent = ''; return; }
      var c = L.classify(prob.value, db.types);
      hint.textContent = prob.value.trim() ? '자동 분류: ' + c.type + (c.hits.length ? ' (키워드 ' + c.hits.join(', ') + ')' : ' — 맞는 키워드가 없어 유형을 골라 주세요') : '';
    }
    prob.addEventListener('input', showAuto); type.addEventListener('change', showAuto); showAuto();
    var errBox = h('div', { 'aria-live': 'polite' });
    function submit(e) {
      e.preventDefault();
      var o = { date: date.value, supplier: sup.value, partNo: pn.value, partName: pname.value, type: type.value, problem: prob.value, qty: qty.value, action: act.value, campus: camp.value };
      var errs = L.checkRecordInput(o, db.types);
      errBox.textContent = '';
      if (errs.length) { errs.forEach(function (x) { errBox.appendChild(h('p', { class: 'alert warn' }, x)); }); return; }
      var rec = L.recordFromInput(o);
      if (orig) {
        ['date', 'supplier', 'partNo', 'partName', 'problem'].forEach(function (k) { orig[k] = rec[k]; });
        ['qty', 'typeManual', 'action', 'campus'].forEach(function (k) { if (rec[k] == null) delete orig[k]; else orig[k] = rec[k]; });
        delete orig.typeIn;
        db.aliases = L.updateAliases(db.aliases, db.records);
        st.editId = ''; save(); render(); toast('수정했습니다.');
        return;
      }
      leaveSample(function () {
        var same = db.records.filter(function (r) { return L.recordKey(r) === L.recordKey(rec); }).length;
        rec.id = 'd' + (++db.seq);
        db.records.push(rec);
        db.aliases = L.updateAliases(db.aliases, db.records);
        db.files.push({ name: '직접 입력', at: nowIso(), added: 1, dup: 0 });
        if (db.files.length > 50) db.files = db.files.slice(-50);
        st.keep = { date: rec.date, supplier: rec.supplier, campus: rec.campus || '' };
        save(); render();
        toast('등록했습니다 (' + rec.supplier + ' · ' + L.typeOf(rec, db.types).type + ').' + (same ? ' 같은 날 같은 내용이 이미 ' + same + '건 있습니다.' : ''));
        var f = document.querySelector('#entryForm input[list="dl-sup"]'); if (f) f.focus();
      });
    }
    return h('section', { class: 'card', id: 'entryCard' },
      h('div', { class: 'card-head' }, h('h2', null, orig ? '불량 수정' : '1. 한 건씩 등록'), orig ? h('span', { class: 'badge warn' }, '수정 중') : null),
      h('form', { id: 'entryForm', onsubmit: submit, novalidate: true },
        datalist('dl-sup', suppliers()),
        h('div', { class: 'form-grid three' },
          field('발생일자 *', date), field('협력사 *', sup, '이미 적은 협력사는 목록에서 고릅니다'), field('불량유형', type),
          field('품번', pn), field('품명', pname), field('수량', qty), field('캠퍼스', camp, '울산 · 인천 어느 쪽 불량인지(선택)')),
        field('내용(문제점) *', prob), hint,
        field('조치', act, '임시 조치·처리 결과(선별·재작업·반품 등). 개선대책은 MASTER PLAN 에서 세웁니다.'),
        errBox,
        h('div', { class: 'btn-row' },
          h('button', { type: 'submit', class: 'btn btn-primary' }, orig ? '수정 저장' : '등록'),
          orig ? h('button', { type: 'button', class: 'btn', onclick: function () { st.editId = ''; render(); } }, '수정 그만두기') : null,
          h('span', { class: 'note' }, '* 표시는 꼭 적어 주세요. 등록하면 발생일자·협력사는 다음 건을 위해 남겨 둡니다.'))));
  }
  function pasteCard() {
    var st = ui.entry;
    var ta = h('textarea', { rows: '6', class: 'mono paste-box', 'aria-label': '엑셀에서 복사한 표', placeholder: '엑셀에서 표를 골라 Ctrl+C(맥은 ⌘+C) → 이 칸을 누르고 Ctrl+V' });
    ta.value = st.paste;
    ta.addEventListener('input', function () { st.paste = ta.value; });
    function read() {
      st.paste = ta.value;
      var p = L.parsePaste(ta.value);
      st.pasted = p; render();
      if (!p.records.length) toast('읽은 줄이 없습니다. 발생일자·협력사·내용 칸이 있는지 확인해 주세요.', true);
    }
    var p = st.pasted, preview = null;
    if (p) {
      var en = L.enrich(p.records, db.types);
      preview = h('div', { class: 'ai-preview' },
        h('h3', null, '읽은 결과 — ' + p.records.length + '건'),
        h('p', { class: 'note' }, (p.header ? '첫 줄의 열 이름으로 칸을 맞췄습니다.' : '열 이름 줄이 없어 「' + L.PASTE_ORDER.map(L.fieldLabel).join(' · ') + '」 순서로 읽었습니다.') +
          (p.skipped ? ' 협력사·내용이 비어 건너뛴 줄 ' + p.skipped + '건.' : '') + (p.badDate ? ' 날짜를 못 읽은 줄 ' + p.badDate + '건(날짜 없이 넣음).' : '')),
        p.records.length ? rawTable(p.records.slice(0, 8)) : null,
        p.records.length > 8 ? h('p', { class: 'note' }, '…외 ' + (p.records.length - 8) + '건') : null,
        en.some(function (r) { return r.type === L.UNCLASSIFIED; }) ? h('p', { class: 'alert warn' }, '유형을 정하지 못한 줄이 ' + en.filter(function (r) { return r.type === L.UNCLASSIFIED; }).length + '건 있습니다. 넣은 뒤 「불량 분석」의 분류 결과에서 고를 수 있습니다.') : null,
        p.records.length ? h('div', { class: 'btn-row' }, h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { addPasted(p.records); } }, p.records.length + '건 넣기 (같은 줄은 건너뜀)')) : null);
    }
    return h('section', { class: 'card' },
      h('h2', null, '2. 엑셀에서 붙여넣기'),
      h('p', { class: 'note' }, '파일을 올리지 않아도 됩니다. 엑셀에서 표 범위를 복사해 아래 칸에 붙여 넣으세요. 첫 줄에 열 이름(발생일자·협력사·품번·내용 등)이 있으면 그 이름으로 맞추고, 없으면 「' + L.PASTE_ORDER.map(L.fieldLabel).join(' · ') + '」 순서로 읽습니다. 불량유형 칸이 비어 있으면 내용 키워드로 나눕니다.'),
      ta,
      h('div', { class: 'btn-row', style: 'margin-top:10px' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: read }, '표 읽기'),
        p ? h('button', { type: 'button', class: 'btn', onclick: function () { st.paste = ''; st.pasted = null; render(); } }, '비우기') : null,
        templateButton()),
      preview);
  }
  function addPasted(recs) {
    leaveSample(function () {
      var m = L.mergeRecords(db.records, recs, db.seq);
      db.records = m.records; db.seq = m.seq;
      db.aliases = L.updateAliases(db.aliases, db.records);
      db.files.push({ name: '붙여넣기', at: nowIso(), added: m.added, dup: m.dup });
      ui.entry.paste = ''; ui.entry.pasted = null;
      save(); render();
      toast('새로 ' + m.added + '건을 넣었습니다' + (m.dup ? ' (같은 줄 ' + m.dup + '건 건너뜀)' : '') + '.');
    });
  }
  function entryList() {
    var st = ui.entry;
    var rows = all();
    var card = h('section', { class: 'card' });
    if (!rows.length) { add(card, [h('h2', null, '3. 등록된 불량'), h('p', { class: 'note' }, '아직 없습니다.')]); return card; }
    var sups = suppliers();
    if (st.supplier && sups.indexOf(st.supplier) < 0) st.supplier = '';
    var sup = select([['', '전체 협력사']].concat(sups.map(function (x) { return [x, x]; })), st.supplier, { 'aria-label': '협력사' });
    sup.addEventListener('change', function () { st.supplier = sup.value; st.page = 0; render(); });
    var ts = select([['', '전체 유형']].concat(L.typeNames(db.types).concat([L.UNCLASSIFIED]).map(function (n) { return [n, n]; })), st.type, { 'aria-label': '불량유형' });
    ts.addEventListener('change', function () { st.type = ts.value; st.page = 0; render(); });
    var q = h('input', { type: 'search', value: st.q, placeholder: '품번·내용·조치 찾기', 'aria-label': '찾기' });
    q.addEventListener('change', function () { st.q = q.value.trim(); st.page = 0; render(); });
    var cs = hasCampus() ? campusSelect(st.campus, function (v) { st.campus = v; st.page = 0; render(); }) : null;
    var shown = rows.filter(function (r) {
      if (st.supplier && r.supplier !== st.supplier) return false;
      if (!L.campusMatch(r, st.campus)) return false;
      if (st.type && r.type !== st.type) return false;
      if (st.q) { var k = st.q.toLowerCase(); if ((r.partNo + ' ' + r.partName + ' ' + r.problem + ' ' + (r.action || '')).toLowerCase().indexOf(k) < 0) return false; }
      return true;
    }).sort(function (a, b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : (+b.id.slice(1) || 0) - (+a.id.slice(1) || 0); });
    var pages = Math.max(1, Math.ceil(shown.length / PAGE));
    if (st.page >= pages) st.page = pages - 1;
    var tb = h('tbody');
    shown.slice(st.page * PAGE, (st.page + 1) * PAGE).forEach(function (r) {
      tb.appendChild(h('tr', { class: (r.type === L.UNCLASSIFIED ? 'row-warn' : '') + (r.id === st.editId ? ' row-edit' : '') },
        h('td', { class: 'nw' }, r.date || '(날짜 없음)'), h('td', { class: 'nw' }, r.campus || ''), h('td', { class: 'nw' }, r.supplier), h('td', { class: 'nw' }, r.partNo), h('td', null, r.partName),
        h('td', { class: 'nw' }, r.type), h('td', { class: 'wide' }, r.problem), h('td', { class: 'num' }, r.qty == null ? '' : fmt(r.qty)), h('td', null, r.action || ''),
        h('td', { class: 'nw' },
          h('button', { type: 'button', class: 'btn btn-sm2', onclick: function () { st.editId = r.id; render(); var c = document.getElementById('entryCard'); if (c) c.scrollIntoView({ block: 'start' }); } }, '수정'), ' ',
          h('button', { type: 'button', class: 'btn btn-sm2 btn-danger', onclick: function () {
            confirmBox('불량 지우기', r.date + ' · ' + r.supplier + ' · ' + r.problem + '\n이 한 건을 지웁니다.', '지우기', function () {
              db.records = db.records.filter(function (x) { return x.id !== r.id; });
              if (st.editId === r.id) st.editId = '';
              save(); render(); toast('지웠습니다.');
            }, true);
          } }, '지우기'))));
    });
    add(card, [
      h('div', { class: 'card-head' }, h('h2', null, '3. 등록된 불량 ' + rows.length + '건'), h('a', { class: 'dp-more', href: '#/analysis' }, '불량 분석', icon('chev'))),
      h('div', { class: 'filters' }, field('협력사', sup), cs ? field('캠퍼스', cs) : null, field('불량유형', ts), field('찾기', q)),
      h('div', { class: 'table-wrap', style: 'margin-top:12px' }, h('table', { class: 'list' },
        h('thead', null, h('tr', null, ['발생일자', '캠퍼스', '협력사', '품번', '품명', '불량유형', '내용', '수량', '조치', ''].map(function (t) { return h('th', null, t); }))), tb)),
      pager(st, pages, shown.length)
    ]);
    return card;
  }
  /* 백업·복원 — JSON(전부: 데이터·사전·Master Plan·확정 이력) · 엑셀(불량 목록) */
  function backupCard() {
    var json = h('input', { type: 'file', accept: '.json,application/json', onchange: function (e) {
      var f = e.target.files[0]; e.target.value = '';
      if (!f) return;
      f.text().then(function (t) {
        var o = JSON.parse(t), d = L.restoreDb(o && o.db ? o.db : o);
        confirmBox('JSON 백업 되살리기', f.name + ' — 불량 ' + d.records.length + '건, Master Plan ' + Object.keys(d.plans).length + '곳.\n지금 데이터를 이 백업으로 바꿉니다.', '되살리기', function () { db = d; save(); toast('되살렸습니다.'); render(); }, true);
      }).catch(function () { toast('JSON 백업 파일을 읽지 못했습니다.', true); });
    } });
    var xl = h('input', { type: 'file', accept: '.xlsx,.xls,.csv', onchange: function (e) {
      var f = e.target.files[0]; e.target.value = '';
      if (!f) return;
      readOne(f).then(function (st) {
        if (st.error) { toast(st.error, true); return; }
        var p = parsedOf(st);
        if (!p || p.r.det.missing.length || !p.r.records.length) { toast('불량 목록을 찾지 못했습니다. 발생일자·협력사·내용 열이 있는지 확인하거나 「파일 업로드」에서 열을 직접 골라 주세요.', true); return; }
        var recs = p.r.records;
        function apply(replace) {
          closeDialog();
          leaveSample(function () {
            if (replace) { db.records = []; db.seq = 0; }
            var m = L.mergeRecords(db.records, recs, db.seq);
            db.records = m.records; db.seq = m.seq;
            db.aliases = L.updateAliases(db.aliases, db.records);
            db.files.push({ name: f.name + ' (엑셀 복원)', at: nowIso(), added: m.added, dup: m.dup });
            save(); render(); toast('엑셀에서 ' + m.added + '건을 읽었습니다' + (m.dup ? ' (같은 줄 ' + m.dup + '건 건너뜀)' : '') + '.');
          });
        }
        openDialog('엑셀에서 복원', h('p', { class: 'pre' }, f.name + ' — 불량 ' + recs.length + '건을 읽었습니다.\n지금 데이터(' + db.records.length + '건)에 더할까요, 지우고 바꿀까요? Master Plan 은 그대로 남습니다.'), [
          h('button', { type: 'button', class: 'btn', onclick: closeDialog }, '그만두기'),
          h('button', { type: 'button', class: 'btn btn-danger', onclick: function () { apply(true); } }, '지우고 바꾸기'),
          h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { apply(false); } }, '더하기')
        ]);
      });
    } });
    return h('section', { class: 'card', id: 'backupCard' }, h('h2', null, '백업 · 복원'),
      h('p', { class: 'note' }, '데이터는 이 브라우저 안에만 있습니다. 브라우저 기록을 지우거나 PC 를 바꾸기 전에 백업을 받아 두세요. JSON 은 불량·사전·Master Plan·확정 이력까지 전부, 엑셀은 불량 목록만(다른 사람에게 보내거나 엑셀에서 고친 뒤 다시 읽을 때) 담습니다. AI 키는 넣지 않습니다.'),
      h('div', { class: 'btn-row' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: function () {
          download('data10-07_백업' + tag() + '_' + today() + '.json', new Blob([JSON.stringify({ app: 'data10-07', at: nowIso(), db: db }, null, 1)], { type: 'application/json' }));
        } }, 'JSON 백업 받기'),
        h('label', { class: 'btn file-btn' }, 'JSON 되살리기', json),
        h('button', { type: 'button', class: 'btn', disabled: !db.records.length, onclick: exportBackupXlsx }, '엑셀 백업 받기'),
        h('label', { class: 'btn file-btn' }, '엑셀에서 복원', xl),
        db.records.length ? h('button', { type: 'button', class: 'btn btn-danger', onclick: clearAll }, '데이터 전부 지우기') : null));
  }

  // ══ 파일 업로드 (선택) ═══════════════════════════════════════
  function decodeCsv(buf) {
    var u8 = new Uint8Array(buf), t;
    try { t = new TextDecoder('utf-8', { fatal: true }).decode(u8); }
    catch (e) { try { t = new TextDecoder('euc-kr').decode(u8); } catch (e2) { t = new TextDecoder('utf-8').decode(u8); } }
    return t.replace(/^﻿/, '');
  }
  function readOne(f, campus) {
    return f.arrayBuffer().then(function (buf) {
      var wb = /\.csv$/i.test(f.name) ? XLSX.read(decodeCsv(buf), { type: 'string' }) : XLSX.read(buf, { type: 'array' });
      var sheets = wb.SheetNames.map(function (n) { return { name: n, aoa: XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null }) }; });
      var pick = L.pickSheet(sheets) || sheets[0];
      var st = { name: f.name, sheets: sheets, sheet: pick ? pick.name : '', ov: null, campus: campus != null ? campus : L.normCampus(L.CAMPUSES.filter(function (c) { return f.name.indexOf(c) >= 0; })[0]) };
      var det = L.detectColumns(pick ? pick.aoa : []);
      var sig = det.headers.map(L.normHead).join('|');
      if (db.mapping && db.mapping[sig]) st.ov = Object.assign({}, db.mapping[sig]);
      return st;
    }).catch(function (e) { return { name: f.name, error: '읽지 못했습니다 (' + (e && e.message || e) + ')' }; });
  }
  /* campus 를 주면(캠퍼스 칸에 올린 파일) 그 캠퍼스로, 아니면 파일 이름에 「울산」「인천」이 있으면 그 캠퍼스로 고릅니다(바꿀 수 있음) */
  function readFiles(fileList, campus) {
    var files = Array.prototype.slice.call(fileList || []);
    if (!files.length) return;
    toast(files.length + '개 파일을 읽는 중…');
    Promise.all(files.map(function (f) { return readOne(f, campus); })).then(function (list) {
      ui.imp = (campus ? (ui.imp || []).filter(function (x) { return x.campus !== campus; }) : []).concat(list); render();
      var c = document.getElementById('importCard'); if (c) c.scrollIntoView({ block: 'start' });
    });
  }
  function dropZone(title, sub, onFiles, multiple) {
    var input = h('input', { type: 'file', accept: '.xlsx,.xls,.csv', multiple: !!multiple, onchange: function (e) { onFiles(e.target.files); e.target.value = ''; } });
    var drop = h('label', { class: 'drop' }, icon('down', 'drop-ico'), h('strong', null, title), h('span', { class: 'note' }, sub), input);
    drop.addEventListener('dragover', function (e) { e.preventDefault(); drop.classList.add('over'); });
    drop.addEventListener('dragleave', function () { drop.classList.remove('over'); });
    drop.addEventListener('drop', function (e) { e.preventDefault(); drop.classList.remove('over'); onFiles(e.dataTransfer.files); });
    return drop;
  }
  /* 캠퍼스별 RAW DATA — 회사는 통합됐지만 시스템은 통합 전이라 울산 · 인천이 따로 뽑은 파일을 각각 올려 합칩니다 */
  function campusCard() {
    var n = L.campusCounts(db.records);
    function last(c) { var f = (db.files || []).filter(function (x) { return x.campus === c; }).slice(-1)[0]; return f ? '마지막으로 올림 ' + f.at.replace('T', ' ').slice(0, 16) + ' · ' + f.name : '아직 올리지 않음'; }
    return h('section', { class: 'card', id: 'campusCard' },
      h('h2', null, '1. 캠퍼스별 공정불량 RAW DATA'),
      h('p', { class: 'note' }, '울산캠퍼스와 인천캠퍼스는 시스템이 달라 RAW DATA 의 열 구성이 다릅니다(울산: 문제점 번호 · 발생 일자 · 협력회사명 · 문제점 상세 … / 인천: QM noti No. · Issued Date · Vendor Name · Problem Description …). 각 칸에 그 캠퍼스 파일을 올리면 열을 알아서 맞추고 캠퍼스를 붙여 합칩니다. 같은 캠퍼스 안의 같은 행은 한 번만 넣고, 캠퍼스가 다르면 같은 내용이어도 따로 셉니다.'),
      h('p', { class: 'note' }, h('strong', null, '다시 올리면 그 캠퍼스만 바뀝니다'), ' — 그 캠퍼스에서 전에 올린 RAW 행을 새 파일로 바꾸고, 다른 캠퍼스와 「불량 등록」에서 직접 적은 행은 그대로 둡니다. 전에 직접 고른 불량유형은 같은 행이면 이어 받습니다.'),
      h('div', { class: 'grid-2 campus-drops' }, L.CAMPUSES.map(function (c) {
        return h('div', { class: 'campus-slot' },
          h('h3', null, L.campusLabel(c), h('span', { class: 'badge ' + (n[c] ? 'ok' : 'muted') }, n[c] + '건')),
          dropZone(L.campusLabel(c) + ' RAW DATA 고르기', '.xlsx · .xls · .csv — 여기로 끌어 놓아도 됩니다', function (fl) { readFiles(fl, c); }, true),
          h('small', { class: 'note' }, last(c)));
      })));
  }
  function parsedOf(st) {
    var sh = st.sheets.filter(function (s) { return s.name === st.sheet; })[0];
    if (!sh) return null;
    return { aoa: sh.aoa, r: L.parseSheet(sh.aoa, st.ov ? { overrides: st.ov } : null) };
  }
  function templateButton() {
    return h('button', { type: 'button', class: 'btn', onclick: function () {
      var aoa = [L.PASTE_ORDER.map(L.fieldLabel), ['2026-10-05', '(예) 협력사 A', '(예) TP-0001', '(예) 브래킷', '용접불량', '(예) 용접부 균열 발생', 1, '(예) 재작업 후 재입고']];
      writeXlsx('공정불량_입력양식.xlsx', { '불량 목록': aoa, '불량유형': [['불량유형 (이 이름 그대로 적어 주세요)']].concat(L.typeNames(db.types).map(function (n) { return [n]; })) },
        { '불량 목록': listMeta(aoa, [12, 14, 14, 14, 12, 40, 8, 30]), '불량유형': { head: [0] } });
    } }, '입력 양식 받기(.xlsx)');
  }
  function renderUpload(main) {
    main.appendChild(pageHead('파일 업로드 (선택)', '이미 Excel·CSV 파일로 정리된 공정불량이 있을 때만 씁니다. 파일이 없으면 「불량 등록」에서 바로 적거나 붙여 넣으면 됩니다. 파일은 서버로 보내지 않고 이 브라우저에서만 읽습니다.'));
    var drop = dropZone('Excel 파일 고르기', '.xlsx · .xls · .csv — 여러 개를 함께 골라도 됩니다(여기로 끌어 놓아도 됨)', function (fl) { readFiles(fl); }, true);
    main.appendChild(campusCard());
    if (ui.imp) main.appendChild(importCard());
    main.appendChild(h('section', { class: 'card' },
      h('h2', null, '2. 그 밖의 파일 올리기 (캠퍼스 구분 없음)'),
      h('p', null, '꼭 필요한 열: ', h('strong', null, '발생일자 · 협력사 · 내용(문제점)'), ' — 품번·품명·불량유형·수량·조치는 있으면 함께 읽습니다. 열 이름이 조금 달라도(업체명·불량내용·P/N·조치사항 등) 자동으로 찾고, 위에 제목 줄이 있어도 머리행을 찾아 읽습니다.'),
      drop,
      h('div', { class: 'btn-row', style: 'margin-top:12px' }, templateButton(), sampleButton(false)),
      h('p', { class: 'note', style: 'margin-top:10px' }, '시험용 가상 파일: 리포의 samples/ 폴더 — 예시_울산캠퍼스_RAW.xlsx · 예시_인천캠퍼스_RAW.xlsx · 예시_협력사_공정불량_RawData.xlsx')));

    var rows = db.records;
    var b = L.dateBounds(rows);
    var cur = h('section', { class: 'card' }, h('h2', null, '지금 들어 있는 데이터'));
    if (!rows.length) cur.appendChild(h('p', { class: 'note' }, '아직 없습니다.'));
    else {
      cur.appendChild(h('div', { class: 'kpis' },
        h('div', { class: 'kpi' }, h('div', { class: 'k' }, '건수'), h('div', { class: 'v' }, fmt(rows.length), h('small', null, '건')), hasCampus() ? h('small', { class: 'note' }, campusText()) : null),
        h('div', { class: 'kpi' }, h('div', { class: 'k' }, '협력사'), h('div', { class: 'v' }, suppliers().length, h('small', null, '곳'))),
        h('div', { class: 'kpi' }, h('div', { class: 'k' }, '기간'), h('div', { class: 'v sm' }, b.min + ' ~ ' + b.max)),
        h('div', { class: 'kpi' }, h('div', { class: 'k' }, '저장 크기'), h('div', { class: 'v sm' }, S.sizeKb() + ' KB'))));
      if (db.files.length) cur.appendChild(h('details', { class: 'files' }, h('summary', null, '가져온 기록 ' + db.files.length + '건'),
        h('ul', null, db.files.slice().reverse().map(function (f) { return h('li', null, f.at.replace('T', ' ').slice(0, 16) + ' · ' + (f.campus ? L.campusLabel(f.campus) + ' · ' : '') + f.name + ' · 새로 ' + f.added + '건' + (f.removed ? ' (전에 올린 ' + f.removed + '건을 바꿈)' : '') + (f.dup ? ', 같은 행 ' + f.dup + '건 건너뜀' : '')); }))));
      var recent = rows.slice().sort(function (a, b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : 0; }).slice(0, 8);
      cur.appendChild(h('h3', { style: 'margin-top:14px' }, '최근 8건'));
      cur.appendChild(rawTable(recent));
      cur.appendChild(h('div', { class: 'btn-row', style: 'margin-top:12px' },
        h('button', { type: 'button', class: 'btn', onclick: exportBackupXlsx }, '엑셀 백업 받기'),
        h('a', { class: 'btn btn-primary', href: '#/analysis' }, '불량 분석으로'),
        h('button', { type: 'button', class: 'btn btn-danger', onclick: clearAll }, '전부 지우기')));
    }
    main.appendChild(cur);
  }
  function rawTable(rows) {
    return h('div', { class: 'table-wrap' }, h('table', { class: 'list' },
      h('thead', null, h('tr', null, ['발생일자', '캠퍼스', '협력사', '품번', '품명', '불량유형', '내용', '수량', '조치'].map(function (t) { return h('th', null, t); }))),
      h('tbody', null, rows.map(function (r) {
        var t = L.typeOf(r, db.types).type;
        return h('tr', { class: t === L.UNCLASSIFIED ? 'row-warn' : '' }, h('td', { class: 'nw' }, r.date || '(날짜 없음)'), h('td', { class: 'nw' }, r.campus || ''), h('td', { class: 'nw' }, r.supplier), h('td', { class: 'nw' }, r.partNo), h('td', null, r.partName),
          h('td', { class: 'nw' }, t), h('td', { class: 'wide' }, r.problem), h('td', { class: 'num' }, r.qty == null ? '' : fmt(r.qty)), h('td', null, r.action || ''));
      }))));
  }
  function importCard() {
    var card = h('section', { class: 'card import', id: 'importCard' }, h('h2', null, '열 연결 확인'));
    var ok = 0, total = 0;
    ui.imp.forEach(function (st, fi) {
      var cs = select([['', '캠퍼스 구분 없음(더하기)']].concat(L.CAMPUSES.map(function (c) { return [c, L.campusLabel(c) + ' RAW(그 캠퍼스만 바꾸기)']; })), st.campus || '', { 'aria-label': '캠퍼스' });
      cs.addEventListener('change', function () { st.campus = cs.value; render(); });
      var box = h('div', { class: 'imp-file' }, h('h3', null, st.campus ? h('span', { class: 'badge ok' }, L.campusLabel(st.campus)) : null, ' ', st.name));
      card.appendChild(box);
      if (st.error) { box.appendChild(h('p', { class: 'alert error' }, st.error)); return; }
      box.appendChild(field('캠퍼스', cs, null, 'inline'));
      if (st.sheets.length > 1) {
        var ss = select(st.sheets.map(function (s) { return s.name; }), st.sheet, { 'aria-label': '시트' });
        ss.addEventListener('change', function () { st.sheet = ss.value; st.ov = null; render(); });
        box.appendChild(field('시트', ss, '머리행이 가장 잘 맞는 시트를 골라 두었습니다.', 'inline'));
      }
      var p = parsedOf(st);
      if (!p) return;
      var det = p.r.det;
      if (det.headerRow < 0) {
        box.appendChild(h('p', { class: 'alert warn' }, '머리행(발생일자·품번·협력사 같은 열 이름 줄)을 찾지 못했습니다. 첫 줄을 머리행으로 보고 열을 직접 골라 주세요.'));
        if (!st.ov) { st.ov = {}; }
        p = { aoa: p.aoa, r: L.parseSheet(p.aoa, { overrides: st.ov, headerRow: 0 }) };
        det = p.r.det;
      }
      var headers = det.headers;
      var grid = h('div', { class: 'map-grid' });
      L.FIELDS.forEach(function (f) {
        var opts = [['-1', '(연결 안 함)']].concat(headers.map(function (x, i) { return [String(i), String.fromCharCode(65 + (i % 26)) + (i >= 26 ? Math.floor(i / 26) : '') + ': ' + (x || '(빈 칸)')]; }));
        var s = select(opts, det.map[f.id] == null ? '-1' : String(det.map[f.id]), { 'aria-label': f.label });
        s.addEventListener('change', function () { st.ov = Object.assign({}, st.ov || {}); st.ov[f.id] = +s.value; render(); });
        grid.appendChild(field(f.label + (f.req ? ' *' : ''), s, null, det.map[f.id] == null && f.req ? 'miss' : ''));
      });
      box.appendChild(grid);
      total++;
      if (det.missing.length) {
        box.appendChild(h('p', { class: 'alert warn' }, '꼭 있어야 하는 열을 연결해 주세요: ' + det.missing.map(function (id) { return L.FIELDS.filter(function (f) { return f.id === id; })[0].label; }).join(', ')));
        return;
      }
      ok++;
      var r = p.r;
      box.appendChild(h('p', { class: 'note' }, '읽은 행 ' + r.records.length + '건' + (r.skipped ? ' · 문제점·협력사가 비어 건너뛴 행 ' + r.skipped + '건' : '') + (r.badDate ? ' · 날짜를 못 읽은 행 ' + r.badDate + '건(날짜 없이 넣고 월별 집계에서는 빠짐)' : '')));
      box.appendChild(rawTable(r.records.slice(0, 5)));
    });
    var modeName = 'impmode';
    var modes = h('div', { class: 'radios' },
      h('label', { class: 'check' }, h('input', { type: 'radio', name: modeName, checked: ui.impMode === 'add', onchange: function () { ui.impMode = 'add'; } }), '지금 데이터에 더하기 (같은 행은 건너뜀)'),
      h('label', { class: 'check' }, h('input', { type: 'radio', name: modeName, checked: ui.impMode === 'replace', onchange: function () { ui.impMode = 'replace'; } }), '지금 데이터를 지우고 바꾸기'));
    card.appendChild(h('h2', { style: 'margin-top:16px' }, '가져오기'));
    var plain = ui.imp.some(function (x) { return !x.error && !x.campus; });
    if (db._sample) card.appendChild(h('p', { class: 'alert info' }, '지금은 예시 데이터라, 가져오면 예시 데이터를 지우고 바꿉니다.'));
    else if (db.records.length && plain) card.appendChild(modes);
    if (ui.imp.some(function (x) { return x.campus; })) card.appendChild(h('p', { class: 'note' }, '캠퍼스를 고른 파일은 그 캠퍼스에서 전에 올린 RAW 행만 바꿉니다(다른 캠퍼스 · 직접 등록한 행은 그대로).'));
    card.appendChild(h('div', { class: 'btn-row' },
      h('button', { type: 'button', class: 'btn btn-primary', disabled: !ok, onclick: doImport }, ok + '/' + total + '개 파일 가져오기'),
      h('button', { type: 'button', class: 'btn', onclick: function () { ui.imp = null; render(); } }, '그만두기')));
    return card;
  }
  /* 캠퍼스를 고른 파일 = 그 캠퍼스 RAW 바꾸기(logic.replaceCampus) · 고르지 않은 파일 = 예전처럼 더하기/바꾸기 */
  function doImport() {
    var recs = [], names = [], byCampus = {};
    db.mapping = db.mapping || {};
    ui.imp.forEach(function (st) {
      if (st.error) return;
      var p = parsedOf(st); if (!p) return;
      var r = st.ov ? L.parseSheet(p.aoa, { overrides: st.ov, headerRow: p.r.det.headerRow < 0 ? 0 : undefined }) : p.r;
      if (r.det.missing.length) return;
      if (st.ov) db.mapping[r.det.headers.map(L.normHead).join('|')] = st.ov;
      if (st.campus) { var b = byCampus[st.campus] || (byCampus[st.campus] = { recs: [], names: [] }); b.recs = b.recs.concat(r.records); b.names.push(st.name); }
      else { recs = recs.concat(r.records); names.push(st.name); }
    });
    var replace = db._sample || !db.records.length || (names.length && ui.impMode === 'replace');
    if (replace) { db.records = []; db.seq = 0; if (db._sample) { db.plans = {}; db.files = []; } delete db._sample; }
    var msg = [];
    L.CAMPUSES.forEach(function (c) {
      var b = byCampus[c]; if (!b) return;
      var m = L.replaceCampus(db.records, c, b.recs, db.seq);
      db.records = m.records; db.seq = m.seq;
      db.files.push({ name: b.names.join(', '), campus: c, at: nowIso(), added: m.added, dup: m.dup, removed: m.removed });
      msg.push(L.campusLabel(c) + ' ' + m.added + '건' + (m.removed ? '(전에 올린 ' + m.removed + '건을 바꿈)' : '') + (m.dup ? ' · 캠퍼스 안 같은 행 ' + m.dup + '건 건너뜀' : ''));
    });
    if (names.length) {
      var m2 = L.mergeRecords(db.records, recs, db.seq);
      db.records = m2.records; db.seq = m2.seq;
      db.files.push({ name: names.join(', '), at: nowIso(), added: m2.added, dup: m2.dup });
      msg.push('새로 ' + m2.added + '건' + (m2.dup ? '(같은 행 ' + m2.dup + '건 건너뜀)' : ''));
    }
    if (db.files.length > 50) db.files = db.files.slice(-50);
    db.aliases = L.updateAliases(db.aliases, db.records);
    save();
    ui.imp = null;
    toast(msg.join(' · ') + ' — 지금 ' + db.records.length + '건' + (hasCampus() ? ' (' + campusText() + ')' : '') + '.');
    location.hash = '#/analysis';
  }
  /* 엑셀 백업 — 붙여넣기·파일 복원으로 그대로 다시 읽히는 순서. 둘째 시트는 자동 분류 결과까지 */
  function exportBackupXlsx() {
    var aoa = L.backupSheet(db.records.slice().sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; }));
    var en = all(), cls = [['발생일자', '협력사', '품번', '불량유형(적용)', '분류 근거', '내용']];
    en.forEach(function (r) { cls.push([r.date, r.supplier, r.partNo, r.type, r.manual ? '직접 지정' : r.column ? '불량유형 칸' : r.hits.join(', '), r.problem]); });
    writeXlsx('공정불량_백업' + tag() + '_' + today() + '.xlsx', { '불량 목록': aoa, '분류 결과': cls },
      { '불량 목록': listMeta(aoa, [12, 14, 14, 14, 12, 40, 8, 30]), '분류 결과': listMeta(cls, [12, 14, 14, 14, 18, 40]) });
    toast('엑셀 백업을 내려받았습니다. 「불량 등록 → 백업·복원」에서 다시 읽을 수 있습니다.');
  }
  function clearAll() {
    confirmBox('전부 지우기', 'Raw data ' + db.records.length + '건과 모든 협력사의 Master Plan 을 지웁니다. 불량유형 사전·규칙·AI 설정은 남습니다.\n되돌릴 수 없으니 필요하면 먼저 설정에서 JSON 백업을 받아 두세요.', '지우기', function () {
      var keep = { types: db.types, depts: db.depts, rules: db.rules, mask: db.mask, mapping: db.mapping, mail: db.mail };
      db = L.emptyDb(); Object.assign(db, keep); save(); toast('지웠습니다.'); render();
    }, true);
  }

  // ══ 불량 분석 ═════════════════════════════════════════════
  function periodControls(state, onChange, bounds) {
    var sup = select([['', '전체 협력사']].concat(suppliers().map(function (s) { return [s, s]; })), state.supplier, { 'aria-label': '협력사' });
    sup.addEventListener('change', function () { state.supplier = sup.value; state.page = 0; onChange(); });
    var f = dateInput(state.from, { 'aria-label': '시작일', min: bounds.min, max: bounds.max });
    var t = dateInput(state.to, { 'aria-label': '종료일', min: bounds.min, max: bounds.max });
    f.addEventListener('change', function () { state.from = f.value; state.page = 0; onChange(); });
    t.addEventListener('change', function () { state.to = t.value; state.page = 0; onChange(); });
    function quick(label, months) {
      return h('button', { type: 'button', class: 'chip-btn', onclick: function () {
        state.to = ''; state.from = months ? L.monthAdd(L.monthOf(bounds.max), -(months - 1)) + '-01' : ''; state.page = 0; onChange();
      } }, label);
    }
    var cs = hasCampus() ? campusSelect(state.campus, function (v) { state.campus = v; state.page = 0; onChange(); }) : null;
    return h('div', { class: 'filters' },
      field('협력사', sup), cs ? field('캠퍼스', cs) : null, field('분석 기간 — 시작', f), field('끝', t),
      h('div', { class: 'field' }, h('span', null, '빠른 기간'), h('div', { class: 'chips' }, quick('최근 3개월', 3), quick('최근 6개월', 6), quick('전체', 0))));
  }
  function renderAnalysis(main) {
    var arg = routeArg();
    if (arg && suppliers().indexOf(arg) >= 0) { ui.an.supplier = arg; history.replaceState(null, '', '#/analysis'); }
    if (!db.records.length) { emptyNotice(main, '불량 분석'); return; }
    var rows = all(), st = ui.an, b = L.dateBounds(rows);
    if (!hasCampus()) st.campus = '';
    var an = L.analyze(rows, { supplier: st.supplier, campus: st.campus, from: st.from, to: st.to }, db.rules);
    main.appendChild(pageHead('불량 분석', '문제점 글의 키워드로 불량유형을 나누고, 협력사별 발생 현황과 개선이 필요한 항목을 찾습니다.'));
    main.appendChild(h('section', { class: 'card' }, periodControls(st, render, b)));
    main.appendChild(h('section', { class: 'dkpis' },
      kpi('불량 발생 건수', 'doc', fmt(an.total), '건', (an.from ? an.from + ' ~ ' + an.to : '') + (hasCampus() && !st.campus ? ' · ' + campusText(an.rows) : st.campus ? ' · ' + (st.campus === '_none' ? '구분 없음' : L.campusLabel(st.campus)) : '')),
      an.hasQty ? kpi('불량 수량', 'layers', fmt(an.qty), '개', '수량 열이 있는 행만 합계') : kpi('협력사', 'factory', an.suppliers.length, '곳'),
      kpi('개선 필요 항목', 'alert', an.issues.length, '건', 'WORST ' + an.worst.length + ' · 동일품번 ' + an.repeatParts.length, 'warn'),
      kpi('유형 미분류', 'list', an.unclassified, '건', an.unclassified ? '아래 표에서 유형을 고르거나 키워드를 더하세요' : '모두 분류됨', an.unclassified ? 'warn' : 'ok')));
    if (!an.total) { main.appendChild(h('section', { class: 'card' }, h('p', null, '이 조건에 맞는 공정불량이 없습니다.'))); return; }

    var charts = [
      panel('불량유형 파레토', null, paretoBars(an.byType)),
      panel('월별 추이', null, h('div', null, h('p', { class: 'dp-sub' }, '최근 12개월까지 · 막대 = 건수'), monthBars(an.months, '월별 불량 건수')))
    ];
    if (!st.supplier) charts.unshift(panel('협력사별 건수', null, hbars(an.suppliers.map(function (s) { return { key: s.name, value: s.count, href: '#/analysis/' + encodeURIComponent(s.name) }; }), { label: '협력사별 불량 건수' })));
    main.appendChild(h('div', { class: st.supplier ? 'grid-2 gap-b' : 'dcharts' }, charts));
    if (!st.supplier && an.suppliers.length > 1) main.appendChild(crossTable(an));

    main.appendChild(issuesCard(an));
    main.appendChild(classifyCard(an));
  }
  /* 개선 필요 항목 — 수강생 기준 두 가지만(2026-10-27): ① 불량유형별 WORST 1·2·3 ② 동일품번 2건 이상.
     불량이 1건이라도 개선 대상이지만, 이 두 경우를 우선 개선 항목으로 올립니다. 근거 글 칸은 두지 않습니다. */
  function issuesCard(an) {
    var r = db.rules, sups = an.suppliers.map(function (x) { return x.name; });
    var ws = {}; an.worst.forEach(function (i) { (ws[i.supplier] || (ws[i.supplier] = []))[i.rank - 1] = i; });
    var cols = []; for (var k = 1; k <= r.worstN; k++) cols.push(k);
    var wt = h('div', { class: 'table-wrap' }, h('table', { class: 'list worst' },
      h('thead', null, h('tr', null, h('th', null, '협력사'), cols.map(function (k) { return h('th', null, 'WORST ' + k); }), h('th', null, 'Master Plan'))),
      h('tbody', null, sups.filter(function (s) { return ws[s]; }).map(function (s) {
        return h('tr', null, h('th', { scope: 'row', class: 'nw' }, s), cols.map(function (k) { var i = ws[s][k - 1]; return h('td', { class: 'nw' }, i ? [i.type, h('small', { class: 'note' }, ' ' + i.count + '건')] : '—'); }),
          h('td', null, h('a', { class: 'btn btn-sm2', href: '#/plan/' + encodeURIComponent(s) }, '계획')));
      }))));
    var pt = an.repeatParts.length ? h('div', { class: 'table-wrap' }, h('table', { class: 'list' },
      h('thead', null, h('tr', null, ['협력사', '품번', '품명', '건수', '불량유형', '최근 발생', 'Master Plan'].map(function (t) { return h('th', null, t); }))),
      h('tbody', null, an.repeatParts.slice(0, 60).map(function (p) {
        return h('tr', null, h('td', { class: 'nw' }, p.supplier), h('td', { class: 'nw' }, p.partNo), h('td', null, p.partName), h('td', { class: 'num' }, p.count), h('td', null, p.types.join(', ')), h('td', { class: 'nw' }, p.last),
          h('td', null, h('a', { class: 'btn btn-sm2', href: '#/plan/' + encodeURIComponent(p.supplier) }, '계획')));
      })))) : h('p', { class: 'note' }, '같은 품번이 ' + r.partRepeatMin + '건 이상인 것이 없습니다.');
    return h('section', { class: 'card', id: 'issuesCard' },
      h('div', { class: 'card-head' }, h('h2', null, '개선 필요 항목 ' + an.issues.length + '건'), h('a', { class: 'dp-more', href: '#/settings' }, '기준 바꾸기', icon('chev'))),
      h('p', { class: 'note' }, '불량이 1건이라도 개선 대상입니다. 그 가운데 아래 두 경우를 우선 개선 항목으로 올리고, Master Plan 초안도 이 항목으로 만듭니다.'),
      h('h3', null, '① 불량유형별 WORST 1·2' + (r.worstN >= 3 ? '·3' : '') + (r.worstN > 3 ? '…' + r.worstN : '') + ' — 협력사마다 건수가 많은 불량유형'),
      an.worst.length ? wt : h('p', { class: 'note' }, '분류된 불량이 없습니다.'),
      h('p', { class: 'note' }, '건수가 같으면 최근에 발생한 유형이 앞입니다. 미분류는 순위에서 뺍니다.'),
      h('h3', { style: 'margin-top:16px' }, '② 동일품번 ' + r.partRepeatMin + '건 이상 — 같은 협력사·같은 품번(불량유형은 달라도 셈) ' + an.repeatParts.length + '개'),
      pt);
  }
  function crossTable(an) {
    var types = an.byType.map(function (t) { return t.type; });
    var m = {}, max = 1;
    an.rows.forEach(function (r) { var k = r.supplier + '\u0000' + r.type; m[k] = (m[k] || 0) + 1; max = Math.max(max, m[k]); });
    var issueKey = {}; an.worst.forEach(function (i) { issueKey[i.supplier + '\u0000' + i.type] = 1; });
    return h('section', { class: 'card' }, h('h2', null, '협력사 × 불량유형'),
      h('p', { class: 'note' }, '칸이 진할수록 건수가 많습니다. 굵은 테두리는 협력사별 WORST 1·2·3 입니다.'),
      h('div', { class: 'table-wrap' }, h('table', { class: 'list cross' },
        h('thead', null, h('tr', null, h('th', null, '협력사'), types.map(function (t) { return h('th', null, t); }), h('th', { class: 'num' }, '합계'))),
        h('tbody', null, an.suppliers.map(function (s) {
          return h('tr', null, h('th', { scope: 'row' }, h('a', { href: '#/analysis/' + encodeURIComponent(s.name) }, s.name)),
            types.map(function (t) {
              var v = m[s.name + '\u0000' + t] || 0, a = v / max;
              return h('td', { class: 'num heat' + (issueKey[s.name + '\u0000' + t] ? ' hot' : ''), style: v ? 'background:rgba(31,95,168,' + (0.08 + a * 0.62).toFixed(2) + ');color:' + (a > 0.55 ? '#fff' : 'inherit') : null }, v || '');
            }), h('td', { class: 'num' }, s.count));
        })))));
  }
  function classifyCard(an) {
    var st = ui.an;
    var names = L.typeNames(db.types);
    var rows = an.rows.filter(function (r) {
      if (st.onlyUnc && r.type !== L.UNCLASSIFIED) return false;
      if (st.type && r.type !== st.type) return false;
      if (st.q) { var q = st.q.toLowerCase(); if ((r.problem + ' ' + r.partNo + ' ' + r.partName).toLowerCase().indexOf(q) < 0) return false; }
      return true;
    }).sort(function (a, b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : 0; });
    var pages = Math.max(1, Math.ceil(rows.length / PAGE));
    if (st.page >= pages) st.page = pages - 1;
    var tsel = select([['', '전체 유형']].concat(names.concat([L.UNCLASSIFIED]).map(function (n) { return [n, n]; })), st.type, { 'aria-label': '불량유형' });
    tsel.addEventListener('change', function () { st.type = tsel.value; st.page = 0; render(); });
    var q = h('input', { type: 'search', value: st.q, placeholder: '문제점·품번 찾기', 'aria-label': '문제점·품번 찾기' });
    q.addEventListener('change', function () { st.q = q.value.trim(); st.page = 0; render(); });
    var unc = h('input', { type: 'checkbox', checked: st.onlyUnc, onchange: function (e) { st.onlyUnc = e.target.checked; st.page = 0; render(); } });
    var tb = h('tbody');
    rows.slice(st.page * PAGE, (st.page + 1) * PAGE).forEach(function (r) {
      var raw = db.records.filter(function (x) { return x.id === r.id; })[0];
      var s = select([['', '(자동) ' + (raw && raw.typeManual ? L.classify(r.problem, db.types).type : r.type)]].concat(names.concat([L.UNCLASSIFIED]).map(function (n) { return [n, n]; })), raw && raw.typeManual || '', { 'aria-label': '불량유형 고르기', class: 'type-sel' + (r.manual ? ' manual' : '') });
      s.addEventListener('change', function () { if (!raw) return; if (s.value) raw.typeManual = s.value; else delete raw.typeManual; save(); render(); });
      tb.appendChild(h('tr', { class: r.type === L.UNCLASSIFIED ? 'row-warn' : '' },
        h('td', { class: 'nw' }, r.date), h('td', { class: 'nw' }, r.supplier), h('td', { class: 'nw' }, r.partNo), h('td', null, r.partName), h('td', { class: 'wide' }, r.problem),
        h('td', null, s), h('td', { class: 'note' }, r.manual ? '직접 지정' : r.column ? '불량유형 칸' : r.hits.join(', '))));
    });
    var kw = h('input', { placeholder: '예: 미첨부', 'aria-label': '더할 키워드' });
    var kwType = select(names, names[0], { 'aria-label': '키워드를 더할 유형' });
    var addKw = h('form', { class: 'kw-add', onsubmit: function (e) {
      e.preventDefault();
      var k = kw.value.trim(); if (!k) return;
      var t = db.types.filter(function (x) { return x.name === kwType.value; })[0];
      if (t.keywords.indexOf(k) < 0) t.keywords.push(k);
      save(); toast('「' + t.name + '」 키워드에 「' + k + '」를 더했습니다.'); render();
    } }, field('키워드', kw), field('유형', kwType), h('button', { type: 'submit', class: 'btn' }, '키워드 더하기'));
    return h('section', { class: 'card' },
      h('div', { class: 'card-head' }, h('h2', null, '분류 결과 ' + rows.length + '건'), h('a', { class: 'dp-more', href: '#/settings' }, '키워드 사전', icon('chev'))),
      h('p', { class: 'note' }, '불량유형은 ① 등록할 때 고른 유형 ② 붙여 넣은 표의 불량유형 칸 ③ 내용에 든 키워드 순서로 정합니다(근거 칸). 틀리면 그 행에서 유형을 직접 고르세요 — 직접 고른 값이 먼저입니다. 같은 낱말이 자주 나오면 키워드로 더하는 편이 낫습니다.'),
      h('div', { class: 'filters' }, field('유형', tsel), field('찾기', q), h('label', { class: 'check' }, unc, '미분류만')),
      h('div', { class: 'table-wrap', style: 'margin-top:12px' }, h('table', { class: 'list' },
        h('thead', null, h('tr', null, ['발생일자', '협력사', '품번', '품명', '문제점', '불량유형', '근거'].map(function (t) { return h('th', null, t); }))), tb)),
      pager(st, pages, rows.length),
      h('details', { class: 'kw-box' }, h('summary', null, '키워드 바로 더하기'), addKw));
  }
  function pager(st, pages, n) {
    if (pages <= 1) return null;
    return h('div', { class: 'pager' },
      h('button', { type: 'button', class: 'btn', disabled: st.page === 0, onclick: function () { st.page--; render(); } }, '이전'),
      h('span', null, (st.page + 1) + ' / ' + pages + ' 쪽 · ' + n + '건'),
      h('button', { type: 'button', class: 'btn', disabled: st.page >= pages - 1, onclick: function () { st.page++; render(); } }, '다음'));
  }

  // ══ MASTER PLAN ═══════════════════════════════════════════
  function planOf(s) { return db.plans[s] || null; }
  function planAnalysis(s, plan) { return L.analyze(all(), { supplier: s, from: plan && plan.from, to: plan && plan.to }, db.rules); }
  function renderPlan(main) {
    var arg = routeArg();
    var sups = suppliers();
    if (arg && sups.indexOf(arg) >= 0) { ui.plan.supplier = arg; ui.plan.parsed = null; history.replaceState(null, '', '#/plan'); }
    if (!db.records.length) { emptyNotice(main, 'MASTER PLAN'); return; }
    if (sups.indexOf(ui.plan.supplier) < 0) {
      var an0 = L.analyze(all(), {}, db.rules);
      ui.plan.supplier = an0.issues.length ? an0.issues[0].supplier : sups[0];
    }
    var s = ui.plan.supplier, plan = planOf(s);
    var b = L.dateBounds(db.records.filter(function (r) { return r.supplier === s; }));
    main.appendChild(pageHead('MASTER PLAN', '협력사 공정불량을 분석해 불량유형별 개선대책·담당부서·추진 일정 초안을 만들고, 간트차트로 계획과 실적을 관리합니다.'));

    // 조건
    var sup = select(sups, s, { 'aria-label': '협력사' });
    sup.addEventListener('change', function () { ui.plan.supplier = sup.value; ui.plan.parsed = null; ui.plan.aiText = ''; render(); });
    var base = dateInput(plan ? plan.baseDate : nextMonday(), { 'aria-label': '계획 착수 기준일' });
    var from = dateInput(plan ? plan.from : '', { 'aria-label': '분석 시작일' });
    var to = dateInput(plan ? plan.to : '', { 'aria-label': '분석 종료일' });
    if (!plan) { from.value = b.min; to.value = b.max; }
    main.appendChild(h('section', { class: 'card plan-ctl' },
      h('div', { class: 'filters' },
        field('협력사', sup), field('계획 착수 기준일', base),
        field('분석기간 — 시작', from), field('끝', to),
        h('div', { class: 'field' }, h('span', { 'aria-hidden': 'true' }, ' '), h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { generate(s, base.value, from.value, to.value); } }, icon('chip'), '분석 및 초안 생성')))));

    if (!plan) {
      var an1 = L.analyze(all(), { supplier: s }, db.rules);
      main.appendChild(h('section', { class: 'card' }, h('h2', null, s + ' — 아직 Master Plan 이 없습니다'),
        h('p', null, '공정불량 ' + an1.total + '건 · 개선 필요 항목 ' + an1.issues.length + '건' + (an1.issues.length ? ' (' + an1.issues.map(function (i) { return i.kind === 'part' ? '동일품번 ' + i.partNo : i.label + ' ' + i.type; }).join(', ') + ')' : '') + '.'),
        h('p', { class: 'note' }, '「분석 및 초안 생성」을 누르면 불량유형 WORST 1·2·3 마다 설정의 기본 대책을, 동일품번 2건 이상인 품번마다 「원인 분석·재발 방지」 한 줄을 넣어 초안을 만듭니다. 간트에서 행을 눌러 직접 고친 뒤 「계획 확정」합니다. Copilot 등으로 다듬을 수도 있습니다(아래 칸).')));
      return;
    }
    var an = planAnalysis(s, plan);
    var stt = L.planStats(plan.rows, today());
    var iss = an.issues.filter(function (i) { return i.supplier === s; });
    main.appendChild(h('section', { class: 'dkpis six' },
      kpi('불량 발생 건수', 'doc', fmt(an.total), '건', an.from + ' ~ ' + an.to),
      kpi('개선 필요 항목', 'alert', iss.length, '건', 'WORST ' + iss.filter(function (i) { return i.kind === 'worst'; }).length + ' · 동일품번 ' + iss.filter(function (i) { return i.kind === 'part'; }).length, 'warn'),
      kpi('개선 활동 수', 'list', stt.total, '건', '추진 ' + stt.days + '일' + (stt.start ? ' (' + stt.start + ' ~ ' + stt.end + ')' : '')),
      kpi('진행중', 'clock', stt.doing, '건', stt.notStarted ? '그 가운데 시작 전 ' + stt.notStarted + '건' : '마감 전 과제', 'doing'),
      kpi('완료', 'check', stt.done, '건', '완료율 ' + stt.rate + '%', 'ok'),
      kpi('미완료 / 보류', 'pause', stt.late + ' / ' + stt.hold, '', stt.late ? '마감이 지났는데 완료가 아님 ' + stt.late + '건' : '미완료 없음', stt.late ? 'bad' : 'hold')));

    main.appendChild(aiSummaryCard(s, plan, an, iss));
    main.appendChild(ganttCard(plan.rows, {
      title: '통합 개선대책 Gantt', state: L.planStateText(plan), scaleKey: ui.plan, showSupplier: false,
      onRow: function (r) { openDrawer(s, r.id); },
      actions: [
        h('button', { type: 'button', class: 'btn', onclick: function () { openDrawer(s, null); } }, icon('plus'), '개선활동 추가'),
        h('button', { type: 'button', class: 'btn btn-ok', onclick: function () { confirmPlanUi(s); } }, icon('tick'), '계획 확정'),
        h('button', { type: 'button', class: 'btn', onclick: function () { reloadDraft(s); } }, icon('redo'), '초안 다시 불러오기'),
        h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { downloadReport([s]); } }, icon('down'), 'Excel 추출 (불량목록 · MASTER PLAN)'),
        h('button', { type: 'button', class: 'btn', onclick: function () { openMail(s); } }, icon('doc'), '메일 만들기 (아웃룩)')
      ]
    }));
    main.appendChild(aiCard(s, plan, an));
  }
  function generate(s, base, from, to) {
    if (!base) { toast('계획 착수 기준일을 골라 주세요.', true); return; }
    if (from && to && to < from) { toast('분석기간의 끝이 시작보다 빠릅니다.', true); return; }
    function go() {
      var old = planOf(s);
      var an = L.analyze(all(), { supplier: s, from: from, to: to }, db.rules);
      var rows = L.buildDraft(an, db.types, { supplier: s, baseDate: base });
      db.plans[s] = { baseDate: base, from: from, to: to, rows: rows, confirmed: old ? old.confirmed : [], aiSummary: '', aiChecks: [], aiAnswer: '' };
      ui.plan.parsed = null; ui.plan.aiText = '';
      save(); render();
      toast(rows.length ? '초안 ' + rows.length + '건을 만들었습니다(기본 대책). 간트에서 행을 눌러 담당자·일정을 고치세요.' : '이 기간에는 초안을 만들 불량유형이 없습니다(미분류만 있음).', !rows.length);
    }
    var p = planOf(s);
    if (p && p.rows.length) confirmBox('초안 다시 만들기', '지금 계획 ' + p.rows.length + '건(실적을 적은 행 ' + p.rows.filter(function (r) { return r.actStart || r.status !== 'open'; }).length + '건 포함)을 지우고 새 초안을 만듭니다.\n확정 이력은 남습니다.', '다시 만들기', go, true);
    else go();
  }
  function reloadDraft(s) {
    var p = planOf(s);
    var fromAi = p && p.aiAnswer;
    confirmBox('초안 다시 불러오기', (fromAi ? '마지막으로 반영한 AI 답으로' : '설정의 기본 대책으로') + ' 초안을 다시 만듭니다.\n지금 계획 ' + p.rows.length + '건은 지워집니다(직접 추가한 행도 포함). 확정 이력은 남습니다.', '다시 불러오기', function () {
      if (fromAi) {
        var r = L.parseAiPlan(p.aiAnswer, { types: db.types, aliases: db.aliases, supplier: s, baseDate: p.baseDate, depts: db.depts });
        p.rows = r.rows; p.aiSummary = r.summary; p.aiChecks = r.checks;
      } else {
        p.rows = L.buildDraft(planAnalysis(s, p), db.types, { supplier: s, baseDate: p.baseDate });
      }
      save(); render(); toast('초안 ' + p.rows.length + '건을 다시 불러왔습니다.');
    }, true);
  }
  function confirmPlanUi(s) {
    var p = planOf(s);
    if (!p.rows.length) { toast('확정할 개선활동이 없습니다.', true); return; }
    var bad = p.rows.filter(function (r) { return !r.planStart || !r.planEnd; }).length;
    var last = L.lastConfirmed(p);
    if (last && !L.planChangedSince(p)) { toast('확정 v' + last.version + ' 이후 계획이 바뀌지 않았습니다.'); return; }
    confirmBox('계획 확정', s + ' Master Plan ' + p.rows.length + '건을 확정 v' + ((p.confirmed || []).length + 1) + '로 남깁니다.' + (bad ? '\n계획일이 비어 있는 활동이 ' + bad + '건 있습니다.' : '') + '\n확정 뒤에도 실적·상태는 계속 적을 수 있고, 계획을 바꾸면 다시 확정하면 됩니다.', '확정', function () {
      var snap = L.confirmPlan(p, nowIso()); save(); render(); toast('확정 v' + snap.version + ' 로 남겼습니다.');
    });
  }
  function aiSummaryCard(s, plan, an, iss) {
    var top = an.byType.filter(function (t) { return t.type !== L.UNCLASSIFIED; })[0];
    var w = iss.filter(function (i) { return i.kind === 'worst'; }).sort(function (a, b) { return a.rank - b.rank; }), pp = iss.filter(function (i) { return i.kind === 'part'; });
    var rule = s + ' 공정불량 ' + an.total + '건 가운데 ' + (top ? top.type + '이(가) ' + top.count + '건(' + top.pct + '%)으로 가장 많습니다.' : '분류된 유형이 없습니다.') +
      (w.length ? ' WORST 는 ' + w.map(function (i) { return i.type; }).join(' · ') + '이고' : '') + (pp.length ? ' 같은 품번이 2건 이상인 품번이 ' + pp.length + '개(' + pp.slice(0, 3).map(function (p) { return p.partNo; }).join(', ') + (pp.length > 3 ? ' 등' : '') + ') 있습니다.' : (w.length ? ' 같은 품번이 반복된 것은 없습니다.' : ''));
    var chips = w.map(function (i) { return h('span', { class: 'chip' }, h('b', null, i.label + ' '), i.type, h('small', null, ' ' + i.count + '건')); })
      .concat(pp.slice(0, 6).map(function (p) { return h('span', { class: 'chip up' }, h('b', null, '동일품번 '), p.partNo, h('small', null, ' ' + p.count + '건')); }));
    return h('section', { class: 'card ai-sum' },
      h('div', { class: 'card-head' }, h('h2', null, icon('chip', 'ai-chip'), '불량 분석 요약')),
      h('div', { class: 'chips' }, chips),
      h('p', null, plan.aiSummary ? h('span', null, plan.aiSummary, h('small', { class: 'note' }, ' — AI 답(검토 필요)')) : rule),
      plan.aiChecks && plan.aiChecks.length ? h('details', null, h('summary', null, 'AI 가 적은 확인 필요 ' + plan.aiChecks.length + '건'), h('ul', null, plan.aiChecks.map(function (c) { return h('li', null, c); }))) : null);
  }

  // 간트 — 행마다 위 얇은 막대 = 계획, 아래 굵은 막대 = 실적(상태 색). 칸 폭을 고정해 머리 칸과 막대가 맞습니다.
  function ganttCard(rows, o) {
    var scale = o.scaleKey.scale;
    var cols = L.ganttColumns(rows, scale, { today: today() });
    var W = scale === 'month' ? 64 : 46;
    var t = today();
    var todayIdx = -1, todayX = -1;
    cols.forEach(function (c, i) { if (c.start <= t && t <= c.end) { todayIdx = i; todayX = i * W + (L.daysBetween(c.start, t) + 0.5) / (L.daysBetween(c.start, c.end) + 1) * W; } });
    // 좁은 화면: 유형·담당부서는 개선대책 칸 안에 작게 넣고 그 칸을 왼쪽에 붙여 둡니다(막대를 옆으로 밀어 봐도 어느 행인지 보이게)
    var narrow = document.documentElement.clientWidth < 700;
    var fixed = narrow ? [['개선대책', 148], ['상태', 74]]
      : o.showSupplier ? [['협력사', 96], ['불량유형', 104], ['개선대책', 240], ['담당부서', 92], ['진행상태', 92]] : [['불량유형', 110], ['개선대책', 260], ['담당부서', 96], ['진행상태', 92]];
    var colgroup = h('colgroup', null, fixed.map(function (f) { return h('col', { style: 'width:' + f[1] + 'px' }); }), cols.map(function () { return h('col', { style: 'width:' + W + 'px' }); }));
    var groups = L.groupSpans(cols);
    var head = h('thead', null,
      h('tr', null, fixed.map(function (f, i) { return h('th', { rowspan: '2', class: 'g-fix' + (i === 0 ? ' g-stick' : '') }, f[0]); }), groups.map(function (g) { return h('th', { colspan: String(g.n), class: 'g-group' }, g.group); })),
      h('tr', null, cols.map(function (c, i) { return h('th', { class: 'g-col' + (i === todayIdx ? ' g-today' : ''), title: c.full }, c.label); })));
    var tb = h('tbody');
    if (!rows.length) tb.appendChild(h('tr', null, h('td', { colspan: String(fixed.length + cols.length), class: 'note' }, '개선활동이 없습니다. 「개선활동 추가」로 넣으세요.')));
    rows.forEach(function (r) {
      var ph = L.phaseOf(r, t);
      var lane = h('div', { class: 'g-lane', style: 'width:' + (cols.length * W) + 'px;background-size:' + W + 'px 100%' });
      var sp = L.span(r.planStart, r.planEnd, cols);
      if (sp) lane.appendChild(h('span', { class: 'g-plan st-' + ph, style: 'left:' + (sp[0] * W + 2) + 'px;width:' + ((sp[1] - sp[0] + 1) * W - 4) + 'px', title: '계획 ' + r.planStart + ' ~ ' + r.planEnd }));
      if (r.actStart) {
        var ae = r.actEnd || (r.status === 'done' ? r.actStart : (t > r.actStart ? t : r.actStart));
        var sa = L.span(r.actStart, ae, cols);
        if (sa) lane.appendChild(h('span', { class: 'g-act st-' + ph + (r.actEnd ? '' : ' ongoing'), style: 'left:' + (sa[0] * W + 2) + 'px;width:' + ((sa[1] - sa[0] + 1) * W - 4) + 'px', title: '실적 ' + r.actStart + ' ~ ' + (r.actEnd || '진행중') }));
      }
      if (todayX >= 0) lane.appendChild(h('span', { class: 'g-now', style: 'left:' + todayX + 'px' }));
      var cells = [];
      if (narrow) cells.push(h('td', { class: 'g-fix g-stick g-act-text' }, h('small', { class: 'note' }, (o.showSupplier ? r.supplier + ' · ' : '') + r.type + ' · ' + r.dept), h('span', { class: 'clip-2' }, r.action)),
        h('td', { class: 'g-fix' }, statusPill(r)));
      else {
      if (o.showSupplier) cells.push(h('td', { class: 'g-fix g-stick' }, r.supplier));
      cells.push(h('td', { class: 'g-fix' }, r.type), h('td', { class: 'g-fix g-act-text' }, h('span', { class: 'clip-2' }, r.action), r.owner ? h('small', { class: 'note' }, r.owner) : null),
        h('td', { class: 'g-fix' }, r.dept), h('td', { class: 'g-fix' }, statusPill(r)));
      }
      tb.appendChild(h('tr', { class: 'clickable', tabindex: '0', 'aria-label': r.type + ' — ' + r.action + ' (눌러서 계획·실적 입력)',
        onclick: function () { o.onRow(r); }, onkeydown: function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); o.onRow(r); } } },
        cells, h('td', { colspan: String(cols.length), class: 'g-lane-cell' }, lane)));
    });
    var tableW = fixed.reduce(function (a, f) { return a + f[1]; }, 0) + cols.length * W;
    var scaleBtns = h('div', { class: 'seg', role: 'group', 'aria-label': '간트 칸' }, [['week', '주차'], ['month', '월']].map(function (x) {
      return h('button', { type: 'button', 'aria-pressed': o.scaleKey.scale === x[0] ? 'true' : 'false', onclick: function () { o.scaleKey.scale = x[0]; render(); } }, x[1]);
    }));
    return h('section', { class: 'card gantt-card' },
      h('div', { class: 'card-head' }, h('h2', null, o.title), o.state ? h('span', { class: 'badge ' + (o.state.indexOf('확정') === 0 ? (o.state.indexOf('수정') > 0 ? 'warn' : 'ok') : 'muted') }, o.state) : null, scaleBtns),
      h('div', { class: 'g-legend' }, legend(), h('span', { class: 'note' }, '위 얇은 막대 = 계획(개선 시작~마감) · 아래 굵은 막대 = 실적 · 세로선 = 오늘 · 행을 누르면 오른쪽 칸에서 계획·실적을 적습니다')),
      h('div', { class: 'table-wrap gantt-wrap' }, h('table', { class: 'gantt', style: 'width:' + tableW + 'px' }, colgroup, head, tb)),
      o.actions ? h('div', { class: 'btn-row', style: 'margin-top:14px' }, o.actions) : null);
  }

  // AI 로 구체화 — 사실만 담은 프롬프트(협력사·품번은 별칭) → 복사 → Copilot(사내 허용)·임시 채팅 ChatGPT·Claude·Gemini 에 붙여 넣기 → 답 붙여 넣기 → 계획 행으로.
  // API 키가 필요 없습니다. 사내 AI 서버로 바로 보내기는 접어 둔 선택 기능입니다.
  function aiCard(s, plan, an) {
    db.aliases = L.updateAliases(db.aliases, db.records);
    var info = L.buildAiPromptInfo({ supplier: s, an: an, rows: plan.rows.filter(function (r) { return r.source === 'template'; }), types: db.types, depts: db.depts, baseDate: plan.baseDate, aliases: db.mask ? db.aliases : null, compact: ui.plan.compact });
    var target = L.AI_TARGETS.filter(function (x) { return x.id === db.aiTarget; })[0] || L.AI_TARGETS[0];
    var tsel = h('div', { class: 'seg ai-targets', role: 'radiogroup', 'aria-label': '붙여 넣을 AI' }, L.AI_TARGETS.map(function (x) {
      return h('button', { type: 'button', role: 'radio', 'aria-checked': x.id === target.id ? 'true' : 'false', onclick: function () { db.aiTarget = x.id; save(); render(); } }, x.name);
    }));
    var compact = h('input', { type: 'checkbox', checked: ui.plan.compact, onchange: function (e) { ui.plan.compact = e.target.checked; render(); } });
    var mask = h('input', { type: 'checkbox', checked: db.mask, onchange: function (e) { db.mask = e.target.checked; save(); render(); } });
    var box = h('textarea', { class: 'prompt-box', readonly: true, 'aria-label': 'AI 프롬프트', rows: '14' }); box.value = info.prompt;
    var ans = h('textarea', { class: 'prompt-box answer', 'aria-label': 'AI 답 붙여 넣기', rows: '10', placeholder: 'AI 답을 여기에 붙여 넣으세요 — 「요약:」 줄, 「| 불량유형 | 개선대책 | 담당부서 | 시작 | 기간 |」 표, 「확인 필요:」 줄을 읽습니다.' });
    ans.value = ui.plan.aiText;
    ans.addEventListener('input', function () { ui.plan.aiText = ans.value; });
    function readAnswer() {
      var r = L.parseAiPlan(ans.value, { types: db.types, aliases: db.aliases, supplier: s, baseDate: plan.baseDate, depts: db.depts });
      ui.plan.parsed = r; ui.plan.aiText = ans.value; render();
      if (!r.rows.length) toast('표 줄을 찾지 못했습니다. 답이 「| 불량유형 | 개선대책 | … |」 표 형식인지 확인해 주세요.', true);
    }
    var send = window.AIPanel.sendButton('사내 AI 서버로 바로 보내기', function () { return info.prompt; }, function (text) { ans.value = text; ui.plan.aiText = text; readAnswer(); }, { toast: toast, settingsHref: '#/settings' });
    var preview = null, p = ui.plan.parsed;
    if (p) {
      var known = L.typeNames(db.types);
      preview = h('div', { class: 'ai-preview' },
        h('h3', null, '읽은 결과 — 개선대책 ' + p.rows.length + '건' + (p.skipped ? ' (읽지 못한 표 줄 ' + p.skipped + '건)' : '')),
        p.summary ? h('p', null, h('strong', null, '요약: '), p.summary) : null,
        p.rows.length ? h('div', { class: 'table-wrap' }, h('table', { class: 'list' },
          h('thead', null, h('tr', null, ['불량유형', '개선대책', '담당부서', '계획 기간'].map(function (x) { return h('th', null, x); }))),
          h('tbody', null, p.rows.map(function (r) { return h('tr', null, h('td', { class: 'nw' }, r.type, known.indexOf(r.type) < 0 ? h('span', { class: 'badge warn', style: 'margin-left:4px' }, '사전에 없음') : null), h('td', { class: 'wide' }, r.action), h('td', { class: 'nw' }, r.dept), h('td', { class: 'nw' }, r.planStart + ' ~ ' + r.planEnd)); })))) : null,
        p.checks.length ? h('div', null, h('strong', null, '확인 필요'), h('ul', null, p.checks.map(function (c) { return h('li', null, c); }))) : null,
        p.rows.length ? h('div', { class: 'radios' },
          h('label', { class: 'check' }, h('input', { type: 'radio', name: 'aimode', checked: ui.plan.mode === 'replace', onchange: function () { ui.plan.mode = 'replace'; } }), '기본 대책·이전 AI 행을 이 답으로 바꾸기 (직접 추가한 행은 남김)'),
          h('label', { class: 'check' }, h('input', { type: 'radio', name: 'aimode', checked: ui.plan.mode === 'append', onchange: function () { ui.plan.mode = 'append'; } }), '지금 계획에 더하기')) : null,
        p.rows.length ? h('div', { class: 'btn-row' }, h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { applyAi(s, p, ans.value); } }, '계획에 반영')) : null);
    }
    var openAi = ui.plan.aiOpen || !!ui.plan.parsed || !!ui.plan.aiText;
    return h('details', { class: 'card ai-card', id: 'aiCard', open: openAi, ontoggle: function (e) { ui.plan.aiOpen = e.target.open; } },
      h('summary', { class: 'card-head' }, h('h2', null, icon('chip', 'ai-chip'), 'AI(Copilot 등)로 개선대책 다듬기 — 선택'), h('span', { class: 'note' }, '프롬프트를 복사해 Copilot 이나 임시 채팅(ChatGPT·Claude·Gemini)에 붙여 넣습니다. API 키가 필요 없습니다.')),
      h('div', { class: 'field' }, h('span', null, '붙여 넣을 AI'), tsel),
      h('ol', { class: 'steps' },
        h('li', null, '「프롬프트 복사」를 누릅니다. 프롬프트에는 이 도구가 계산한 사실만 들어 있고, 협력사·품번은 별칭으로 가립니다.'),
        h('li', null, target.how + (target.id === 'copilot' ? '' : ' 임시 채팅은 대화 기록이 남지 않습니다. 회사 규정이 허용하는 범위에서만 쓰세요.')),
        h('li', null, 'AI 답 전체를 복사해 아래 「AI 답」 칸에 붙여 넣고 「답 읽기」 → 「계획에 반영」 → 간트에서 담당자·일정을 다듬은 뒤 「계획 확정」합니다.')),
      h('label', { class: 'check' }, mask, '협력사 이름과 품번을 별칭(업체1, 품목1)으로 가리기 — 권장(특히 임시 채팅). 대응표는 이 브라우저에만 있고, AI 답의 별칭은 원래 이름으로 되돌립니다.'),
      h('label', { class: 'check' }, compact, '짧게 — 입력 글자 수가 빠듯할 때(문제점 기록을 유형마다 2건으로, 지금 초안 줄을 뺌)'),
      info.leaks.length ? h('p', { class: 'alert warn' }, '가리지 못한 이름이 프롬프트에 남아 있습니다: ' + info.leaks.join(', ') + ' — 붙여 넣기 전에 직접 지워 주세요.') : null,
      h('details', { class: 'prompt-details', open: ui.plan.showPrompt, ontoggle: function (e) { ui.plan.showPrompt = e.target.open; } }, h('summary', null, '프롬프트 보기 (' + info.prompt.length + '자)'), box),
      h('div', { class: 'btn-row' }, h('button', { type: 'button', class: 'btn btn-primary', id: 'copyPrompt', onclick: function () { copyText(info.prompt); } }, '프롬프트 복사 → ' + target.name.replace(/ \(.*\)$/, '') + ' 에 붙여 넣기')),
      h('details', { class: 'prompt-details' }, h('summary', null, '(선택) 사내 AI 서버로 바로 보내기 — 설정에 OpenAI 호환 주소를 적은 경우만'), h('div', { class: 'btn-row' }, send)),
      field('AI 답', ans),
      h('div', { class: 'btn-row' }, h('button', { type: 'button', class: 'btn', onclick: readAnswer }, '답 읽기'), ui.plan.parsed ? h('button', { type: 'button', class: 'btn', onclick: function () { ui.plan.parsed = null; ui.plan.aiText = ''; render(); } }, '답 지우기') : null),
      preview);
  }
  function applyAi(s, p, text) {
    var plan = planOf(s);
    if (ui.plan.mode === 'replace') plan.rows = plan.rows.filter(function (r) { return r.source === 'manual'; }).concat(p.rows);
    else plan.rows = plan.rows.concat(p.rows);
    plan.aiSummary = p.summary; plan.aiChecks = p.checks; plan.aiAnswer = text;
    ui.plan.parsed = null; ui.plan.aiText = '';
    save(); render();
    toast('AI 개선대책 ' + p.rows.length + '건을 계획에 넣었습니다. 행을 눌러 담당자와 일정을 확인하세요.');
  }

  // ── 오른쪽 칸: 계획·실적 입력 ─────────────────────────────
  var drawer = document.getElementById('drawer'), drawerBody = document.getElementById('drawerBody'), drawerBack = document.getElementById('drawerBackdrop'), drawerLast = null;
  function closeDrawer(silent) {
    if (drawer.hidden) return;
    drawer.hidden = true; drawerBack.hidden = true; document.body.classList.remove('drawer-open');
    if (!silent && drawerLast && document.body.contains(drawerLast)) drawerLast.focus();
  }
  document.getElementById('drawerClose').addEventListener('click', function () { closeDrawer(); });
  drawerBack.addEventListener('click', function () { closeDrawer(); });
  function openDrawer(s, id) {
    var plan = planOf(s);
    var orig = id ? plan.rows.filter(function (r) { return r.id === id; })[0] : null;
    var r = orig ? JSON.parse(JSON.stringify(orig)) : L.makeRow({ supplier: s, type: (plan.rows[0] || {}).type || '', action: '', dept: db.depts[0] || '품질', planStart: plan.baseDate, planEnd: L.addDays(plan.baseDate, 13), source: 'manual' });
    drawerLast = document.activeElement;
    document.getElementById('drawerTitle').textContent = orig ? '개선활동 — 계획·실적 입력' : '개선활동 추가';
    drawerBody.textContent = '';
    var typeIn = h('input', { value: r.type, list: 'dl-types' });
    var act = h('textarea', { rows: '3' }); act.value = r.action;
    var dept = h('input', { value: r.dept, list: 'dl-depts' });
    var owner = h('input', { value: r.owner, placeholder: '예: 홍길동 책임' });
    var ps = dateInput(r.planStart), pe = dateInput(r.planEnd), as = dateInput(r.actStart), ae = dateInput(r.actEnd);
    var phaseNote = h('small', { class: 'hint' });
    var note = h('textarea', { rows: '2' }); note.value = r.note;
    var status = r.status;
    var seg = h('div', { class: 'status-seg', role: 'radiogroup', 'aria-label': '실적 상태' });
    function drawSeg() {
      seg.textContent = '';
      ['open', 'done', 'hold'].forEach(function (k) {
        seg.appendChild(h('button', { type: 'button', role: 'radio', 'aria-checked': status === k ? 'true' : 'false', class: 'st-btn st-' + k, onclick: function () {
          status = k;
          if (k === 'done' && !ae.value) { ae.value = today(); if (!as.value) as.value = ps.value && ps.value <= today() ? ps.value : today(); }
          drawSeg();
        } }, k === 'open' ? '진행중 · 미완료' : L.STATUS[k]));
      });
      var ph = L.phaseOf({ status: status, planStart: ps.value, planEnd: pe.value, actStart: as.value }, today());
      phaseNote.textContent = '지금 진행상태: ' + L.PHASE[ph] + (status === 'open' ? ' — 완료·보류가 아니면 날짜로 정해집니다(개선 마감 전 = 진행중, 마감이 지나면 미완료)' : '');
    }
    drawSeg();
    [ps, pe, as].forEach(function (x) { x.addEventListener('change', drawSeg); });
    var errBox = h('div', { 'aria-live': 'polite' });
    function collect() {
      delete r.progress;
      return Object.assign(r, { type: typeIn.value.trim() || '기타', action: act.value.trim(), dept: dept.value.trim() || db.depts[0] || '품질', owner: owner.value.trim(),
        planStart: ps.value, planEnd: pe.value, actStart: as.value, actEnd: ae.value, status: status, note: note.value.trim() });
    }
    function doSave() {
      var x = collect(), errs = L.checkPlanRow(x);
      errBox.textContent = '';
      if (errs.length) { errs.forEach(function (e) { errBox.appendChild(h('p', { class: 'alert warn' }, e)); }); return; }
      if (orig) { delete orig.progress; Object.assign(orig, x); } else plan.rows.push(x);
      save(); closeDrawer(); render(); toast(orig ? '저장했습니다.' : '개선활동을 더했습니다.');
    }
    add(drawerBody, [
      datalist('dl-types', L.typeNames(db.types)), datalist('dl-depts', db.depts),
      h('p', { class: 'note' }, s + (orig && orig.source !== 'manual' ? ' · ' + L.sourceText(orig) + '에서 온 행' : '') + (orig && orig.replyAt ? ' · 협력사 회신 반영 ' + orig.replyAt.replace('T', ' ').slice(0, 16) : '')),
      h('h3', null, '계획'),
      field('불량유형', typeIn), field('개선대책', act), h('div', { class: 'form-grid' }, field('담당부서(협력사 내부)', dept), field('담당자', owner), field('개선 시작(계획)', ps), field('개선 마감(계획)', pe)),
      h('h3', null, '실적'),
      h('div', { class: 'form-grid' }, field('실적 시작', as), field('실적 완료', ae)),
      h('div', { class: 'field' }, h('span', null, '진행상태'), seg, phaseNote),
      field('비고(근거·메모)', note),
      errBox,
      h('div', { class: 'btn-row drawer-actions' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: doSave }, '저장'),
        orig ? h('button', { type: 'button', class: 'btn btn-danger', onclick: function () {
          plan.rows = plan.rows.filter(function (x) { return x.id !== orig.id; }); save(); closeDrawer(); render(); toast('지웠습니다.');
        } }, '이 활동 지우기') : null,
        h('button', { type: 'button', class: 'btn', onclick: function () { closeDrawer(); } }, '닫기'))
    ]);
    drawer.hidden = false; drawerBack.hidden = false; document.body.classList.add('drawer-open');
    drawer.scrollTop = 0;
    setTimeout(function () { (orig ? owner : act).focus(); }, 30);
  }

  // ══ 추진현황 ═══════════════════════════════════════════════
  function renderProgress(main) {
    main.appendChild(pageHead('추진현황', '모든 협력사의 개선과제를 진행상태(진행중 · 완료 · 미완료 · 보류)로 봅니다. 행을 누르면 실적과 상태를 적을 수 있습니다.'));
    var sups = Object.keys(db.plans).filter(function (s) { return db.plans[s].rows.length; }).sort();
    var arg = routeArg();
    if (arg && L.PHASE[arg]) { ui.prog.phase = arg; history.replaceState(null, '', '#/progress'); }
    if (!sups.length) {
      main.appendChild(h('section', { class: 'card' }, h('p', null, '아직 Master Plan 이 없습니다.'), h('a', { class: 'btn btn-primary', href: '#/plan' }, 'MASTER PLAN 으로')));
      return;
    }
    var st = ui.prog, t = today();
    if (st.supplier && sups.indexOf(st.supplier) < 0) st.supplier = '';
    var sup = select([['', '전체 협력사']].concat(sups.map(function (s) { return [s, s]; })), st.supplier, { 'aria-label': '협력사' });
    sup.addEventListener('change', function () { st.supplier = sup.value; render(); });
    var ss = select([['', '전체 진행상태']].concat(L.PHASE_ORDER.map(function (k) { return [k, L.PHASE[k]]; })), st.phase, { 'aria-label': '진행상태' });
    ss.addEventListener('change', function () { st.phase = ss.value; render(); });
    main.appendChild(h('section', { class: 'card' }, h('div', { class: 'filters' }, field('협력사', sup), field('진행상태', ss))));

    var rows = [];
    sups.forEach(function (s) { if (!st.supplier || st.supplier === s) rows = rows.concat(db.plans[s].rows); });
    var ps = L.planStats(rows, t);
    main.appendChild(panel('개선 MASTER PLAN 현황 — 기준일 ' + t, null, phaseTiles(ps), 'p-tiles'));

    var tb = h('tbody');
    sups.forEach(function (s) {
      var p = db.plans[s], x = L.planStats(p.rows, t);
      tb.appendChild(h('tr', null, h('th', { scope: 'row' }, h('a', { href: '#/plan/' + encodeURIComponent(s) }, s)), h('td', null, L.planStateText(p)), h('td', { class: 'stack-cell' }, stackBar(x)),
        h('td', { class: 'num' }, x.total), h('td', { class: 'num' }, x.doing), h('td', { class: 'num' }, x.done), h('td', { class: 'num' }, x.late ? h('span', { class: 'late' }, String(x.late)) : '0'), h('td', { class: 'num' }, x.hold), h('td', { class: 'num' }, x.rate + '%')));
    });
    main.appendChild(h('section', { class: 'card' }, h('h2', null, '협력사별'), h('div', { class: 'table-wrap' }, h('table', { class: 'list' },
      h('thead', null, h('tr', null, ['협력사', '계획 상태', '진행', '전체', '진행중', '완료', '미완료', '보류', '완료율'].map(function (x) { return h('th', null, x); }))), tb)), legend()));

    var shown = rows.filter(function (r) { return !st.phase || L.phaseOf(r, t) === st.phase; })
      .sort(function (a, b) { return (a.planEnd || '9') < (b.planEnd || '9') ? -1 : (a.planEnd || '9') > (b.planEnd || '9') ? 1 : 0; });
    main.appendChild(ganttCard(shown, { title: '개선활동 ' + shown.length + '건' + (st.phase ? ' (' + L.PHASE[st.phase] + ')' : ''), scaleKey: st, showSupplier: true, onRow: function (r) { openDrawer(r.supplier, r.id); } }));
  }

  // ══ Excel 추출 (보고서 출력) — 「불량목록」·「MASTER PLAN」 두 시트 ═════════════════════
  function reportCtx(list, campus) { return { suppliers: list, records: all(), plans: db.plans, rules: db.rules, today: today(), types: db.types, depts: db.depts, campus: campus || '' }; }
  function reportName(list, campus) { return 'MasterPlan_' + (list.length === 1 ? L.safeFileName(list[0]) : '전체') + (campus && campus !== '_none' ? '_' + campus : '') + tag() + '_' + today() + '.xlsx'; }
  function downloadReport(list, campus) {
    var book = L.reportBook(reportCtx(list, campus));
    var name = reportName(list, campus);
    writeXlsx(name, book.sheets, book.meta);
    toast(name + ' 을 내려받았습니다. 협력사에 보내고, 회신 파일은 「진행사항 업로드」에서 올립니다.');
  }
  function renderReport(main) {
    if (!db.records.length) { emptyNotice(main, '보고서 출력'); return; }
    main.appendChild(pageHead('보고서 출력 · Excel 추출', 'Excel 은 「불량목록」과 「MASTER PLAN」 두 시트만 담습니다. MASTER PLAN 시트는 협력사가 노란 칸을 고쳐 회신하는 양식이고, 회신 파일은 「진행사항 업로드」에서 다시 올립니다.'));
    var sups = suppliers();
    var st = ui.rep;
    if (st.supplier && sups.indexOf(st.supplier) < 0) st.supplier = '';
    var sel = select([['', '전체 (' + sups.length + '곳)']].concat(sups.map(function (s) { return [s, s]; })), st.supplier, { 'aria-label': '보고서 협력사' });
    sel.addEventListener('change', function () { st.supplier = sel.value; render(); });
    var list = st.supplier ? [st.supplier] : sups;
    if (!hasCampus()) st.campus = '';
    var cs = hasCampus() ? campusSelect(st.campus, function (v) { st.campus = v; render(); }, '전체 캠퍼스 (울산 + 인천)') : null;
    var sheets = L.reportSheets(reportCtx(list, st.campus));
    var noPlan = list.filter(function (s) { return !db.plans[s] || !db.plans[s].rows.length; });
    var unconfirmed = list.filter(function (s) { return db.plans[s] && db.plans[s].rows.length && (!L.lastConfirmed(db.plans[s]) || L.planChangedSince(db.plans[s])); });
    var nPlan = sheets[L.SHEET_PLAN].length - 5;
    main.appendChild(h('section', { class: 'card' },
      h('div', { class: 'filters' }, field('협력사', sel, '협력사에 보낼 때는 그 협력사 하나를 고르세요'), cs ? field('캠퍼스', cs, '불량목록 시트에만 적용(MASTER PLAN 은 협력사 단위)') : null),
      unconfirmed.length ? h('p', { class: 'alert warn' }, '확정하지 않았거나 확정 뒤 계획을 바꾼 협력사: ' + unconfirmed.join(', ') + ' — MASTER PLAN 시트 위쪽 「계획 상태」에 그대로 적힙니다.') : null,
      noPlan.length ? h('p', { class: 'alert info' }, 'Master Plan 이 아직 없는 협력사: ' + noPlan.join(', ') + ' — 불량목록에는 들어가고 MASTER PLAN 에는 빠집니다.') : null,
      h('div', { class: 'table-wrap' }, h('table', { class: 'list sheets' }, h('thead', null, h('tr', null, h('th', null, '시트'), h('th', null, '내용'), h('th', { class: 'num' }, '행'))),
        h('tbody', null,
          h('tr', null, h('td', { class: 'nw' }, L.SHEET_LIST), h('td', null, '캠퍼스 · 발생일자 · 협력사 · 품번 · 품명 · 불량유형 · 내용 · 수량 · 조치 · 관리번호 (거르기 · 틀 고정)'), h('td', { class: 'num' }, sheets[L.SHEET_LIST].length - 1)),
          h('tr', null, h('td', { class: 'nw' }, L.SHEET_PLAN), h('td', null, '현황(전체 · 진행중 · 완료 · 미완료 · 보류) + 개선대책 · 담당부서 · 담당자 · 개선 시작/마감 · 실적 · 진행상태 · 비고. 진행상태 칸은 엑셀 조건부 서식(완료 초록 · 진행중 주황 · 미완료 빨강 · 보류 회색)이라 협력사가 값을 바꾸면 색도 바뀝니다. 협력사가 고치는 칸은 노란 머리, 진행상태·담당부서는 고르기 목록. 숨긴 ID·기준 칸으로 회신을 맞춥니다'), h('td', { class: 'num' }, Math.max(0, nPlan)))))),
      h('div', { class: 'btn-row', style: 'margin-top:14px' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { downloadReport(list, st.campus); } }, icon('down'), 'Excel 추출 (.xlsx)'),
        h('button', { type: 'button', class: 'btn', id: 'mailBtn', disabled: list.length !== 1, title: list.length !== 1 ? '협력사 하나를 고르면 쓸 수 있습니다' : null, onclick: function () { openMail(list[0]); } }, icon('doc'), '메일 만들기 (아웃룩)'),
        h('a', { class: 'btn', href: '#/reply' }, icon('redo'), '회신 파일 올리기'),
        h('button', { type: 'button', class: 'btn', onclick: function () { window.print(); } }, '인쇄 / PDF'))));
    // 미리보기(인쇄에도 씀)
    var pv = h('section', { class: 'card report-preview' }, h('h2', null, '협력사 공정불량 개선 MASTER PLAN'), h('p', { class: 'note' }, '기준일 ' + today()));
    list.forEach(function (s) {
      var p = db.plans[s] || { rows: [], confirmed: [] }, an = planAnalysis(s, db.plans[s]), x = L.planStats(p.rows, today());
      if (!p.rows.length) { pv.appendChild(h('h3', null, s + ' — Master Plan 없음')); pv.appendChild(h('p', null, '분석 기간 ' + an.from + ' ~ ' + an.to + ' · 공정불량 ' + an.total + '건 · 개선 필요 항목 ' + an.issues.length + '건')); return; }
      pv.appendChild(h('h3', null, s + ' — ' + L.planStateText(p)));
      pv.appendChild(h('p', null, '분석 기간 ' + an.from + ' ~ ' + an.to + ' · 공정불량 ' + an.total + '건 · 개선 필요 항목 ' + an.issues.length + '건 · 개선과제 ' + x.total + '건(진행중 ' + x.doing + ' · 완료 ' + x.done + ' · 미완료 ' + x.late + ' · 보류 ' + x.hold + ')'));
      if (p.aiSummary) pv.appendChild(h('p', { class: 'note' }, 'AI 요약(검토 필요): ' + p.aiSummary));
      pv.appendChild(h('div', { class: 'table-wrap' }, h('table', { class: 'list' },
        h('thead', null, h('tr', null, ['No', '불량유형', '개선대책', '담당부서', '담당자', '개선 시작~마감', '실적', '진행상태'].map(function (c) { return h('th', null, c); }))),
        h('tbody', null, p.rows.map(function (r, i) {
          return h('tr', null, h('td', { class: 'num' }, i + 1), h('td', { class: 'nw' }, r.type), h('td', { class: 'wide' }, r.action), h('td', { class: 'nw' }, r.dept), h('td', null, r.owner),
            h('td', { class: 'nw' }, (r.planStart || '') + ' ~ ' + (r.planEnd || '')), h('td', { class: 'nw' }, r.actStart ? r.actStart + ' ~ ' + (r.actEnd || '') : ''), h('td', null, statusPill(r)));
        })))));
    });
    main.appendChild(pv);
  }

  // ══ 진행사항 업로드 — 협력사가 고쳐 회신한 MASTER PLAN 엑셀 → 바뀐 칸만 반영 · 이력 ═════════
  function renderReply(main) {
    var st = ui.reply;
    main.appendChild(pageHead('진행사항 업로드', '협력사가 고쳐 회신한 Excel(이 도구에서 추출한 MASTER PLAN 시트)을 올리면, 숨긴 ID 칸으로 행을 맞춰 바뀐 칸만 미리 보여 주고 반영합니다. 파일은 서버로 보내지 않고 이 브라우저에서만 읽습니다.'));
    var input = h('input', { type: 'file', accept: '.xlsx,.xls', id: 'replyFile', onchange: function (e) { var f = e.target.files[0]; e.target.value = ''; if (f) readReply(f); } });
    var drop = h('label', { class: 'drop' }, icon('down', 'drop-ico'), h('strong', null, '회신 Excel 고르기'), h('span', { class: 'note' }, '.xlsx — 여기로 끌어 놓아도 됩니다'), input);
    drop.addEventListener('dragover', function (e) { e.preventDefault(); drop.classList.add('over'); });
    drop.addEventListener('dragleave', function () { drop.classList.remove('over'); });
    drop.addEventListener('drop', function (e) { e.preventDefault(); drop.classList.remove('over'); var f = e.dataTransfer.files[0]; if (f) readReply(f); });
    main.appendChild(h('section', { class: 'card' },
      h('h2', null, '1. 회신 파일 올리기'),
      h('ol', { class: 'steps' },
        h('li', null, '「보고서 출력 · Excel 추출」에서 협력사 하나를 골라 내려받거나 「메일 만들기」로 아웃룩 메일 초안을 만들어 협력사에 보냅니다.'),
        h('li', null, '협력사는 MASTER PLAN 시트의 노란 칸(개선대책 · 담당부서 · 담당자 · 개선 시작/마감 · 실적 · 진행상태 · 비고)을 고쳐 회신합니다. 새 활동은 표 아래 빈 줄에 ID 없이 적습니다. 3차에 보낸 파일(진행률 칸 · 진행 중 · 지연 · 예정)도 그대로 읽습니다.'),
        h('li', null, '회신 파일을 여기에 올리고, 바뀐 내용을 확인한 뒤 「반영」합니다. 같은 파일을 다시 올려도 두 번 반영되지 않습니다.')),
      drop));
    if (st.error) main.appendChild(h('p', { class: 'alert warn' }, st.error));
    if (st.diff) main.appendChild(replyPreview());
    main.appendChild(replyHistory());
  }
  function readReply(f) {
    var st = ui.reply;
    f.arrayBuffer().then(function (buf) {
      var u8 = new Uint8Array(buf), wb = XLSX.read(u8, { type: 'array' });
      var name = wb.SheetNames.indexOf(L.SHEET_PLAN) >= 0 ? L.SHEET_PLAN : null, parsed = null;
      (name ? [name] : wb.SheetNames).some(function (n) {
        var p = L.parseProgressSheet(XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: null }), db.types);
        if (p.ok) { parsed = p; parsed.sheet = n; return true; }
        if (!parsed) parsed = p;
        return false;
      });
      st.file = { name: f.name, size: f.size, hash: L.fingerprint(u8) };
      st.take = {};
      if (!parsed || !parsed.ok) { st.parsed = null; st.diff = null; st.error = f.name + ' — ' + (parsed ? parsed.error : '시트를 읽지 못했습니다.'); render(); return; }
      st.error = ''; st.parsed = parsed;
      st.diff = L.diffProgress(db.plans, parsed, { types: db.types, knownSuppliers: suppliers() });
      render();
      var c = st.diff.counts;
      toast(c.changed + c.new ? '바뀐 행 ' + c.changed + '건 · 새 활동 ' + c.new + '건을 찾았습니다. 확인 후 「반영」하세요.' : '이 파일에는 반영할 것이 없습니다(이미 반영했거나 바뀐 칸 없음).');
    }).catch(function (e) { st.error = f.name + ' — 읽지 못했습니다 (' + (e && e.message || e) + ')'; st.diff = null; render(); });
  }
  function replyPreview() {
    var st = ui.reply, d = st.diff, c = d.counts;
    var by = h('input', { value: db.replyBy || '', placeholder: '예: 품질담당 이준호', autocomplete: 'off', 'aria-label': '반영한 사람' });
    by.addEventListener('change', function () { db.replyBy = by.value.trim(); save(); });
    var prev = (db.updates || []).filter(function (u) { return u.hash === st.file.hash; }).slice(-1)[0];
    function chipN(n, label, cls) { return h('span', { class: 'chip ' + (cls || '') }, label + ' ', h('b', null, String(n))); }
    var tb = h('tbody');
    d.items.forEach(function (it) {
      if (it.kind === 'same') return;
      var body;
      if (it.kind === 'changed') {
        body = h('ul', { class: 'diff' }, it.changes.map(function (x) { return h('li', null, h('b', null, x.label), ' ', h('del', null, String(x.show[0] === '' ? '(빈칸)' : x.show[0])), ' → ', h('ins', null, String(x.show[1] === '' ? '(빈칸)' : x.show[1]))); }),
          it.conflicts.map(function (x) {
            var cb = h('input', { type: 'checkbox', checked: !!(st.take[it.line] && st.take[it.line][x.k]), onchange: function (e) { st.take[it.line] = st.take[it.line] || {}; st.take[it.line][x.k] = e.target.checked; } });
            return h('li', { class: 'conflict' }, h('b', null, x.label + ' — 충돌'), ' 이 도구: ', h('span', null, String(x.show[0] || '(빈칸)')), ' · 회신: ', h('span', null, String(x.show[1] || '(빈칸)')),
              h('label', { class: 'check' }, cb, '회신 값으로'), h('small', { class: 'note' }, ' (내보낸 뒤 이 도구에서도 고친 칸입니다. 고르지 않으면 이 도구 값을 둡니다)'));
          }));
      } else if (it.kind === 'new') body = h('span', null, '새 활동 — ' + [it.vals.type, it.vals.dept, (it.vals.planStart || '') + '~' + (it.vals.planEnd || '')].filter(Boolean).join(' · '));
      else body = h('span', { class: it.kind === 'invalid' ? 'late' : 'note' }, it.reasons.join(' · '));
      var badge = { changed: ['바뀜', 'ok'], new: ['새 활동', 'ok'], unknown: ['모르는 행', 'muted'], invalid: ['확인 필요', 'warn'] }[it.kind];
      tb.appendChild(h('tr', { class: 'reply-' + it.kind }, h('td', { class: 'num' }, it.line), h('td', { class: 'nw' }, h('span', { class: 'badge ' + badge[1] }, badge[0])), h('td', { class: 'nw' }, it.supplier || ''),
        h('td', { class: 'wide' }, it.action), h('td', null, body)));
    });
    var applyN = c.changed + c.new;
    return h('section', { class: 'card', id: 'replyPreview' },
      h('div', { class: 'card-head' }, h('h2', null, '2. 바뀐 내용 확인 — ' + st.file.name), st.parsed.sheet ? h('span', { class: 'badge muted' }, '시트 ' + st.parsed.sheet) : null),
      prev ? h('p', { class: 'alert info' }, '이 파일은 ' + prev.at.replace('T', ' ').slice(0, 16) + ' 에 이미 반영한 파일과 같습니다. 그 뒤 바뀐 것만 다시 계산했습니다.') : null,
      !st.parsed.hasId ? h('p', { class: 'alert warn' }, '숨긴 ID 칸이 없는 파일입니다. 협력사·개선대책 글로만 행을 맞춥니다.') : null,
      h('div', { class: 'chips' }, chipN(c.changed, '바뀐 행'), chipN(c.fields, '바뀐 칸'), chipN(c.new, '새 활동'), chipN(c.same, '그대로', 'muted'), chipN(c.conflicts, '충돌', c.conflicts ? 'up' : 'muted'), chipN(c.invalid, '확인 필요', c.invalid ? 'up' : 'muted'), chipN(c.unknown, '모르는 행', 'muted')),
      tb.children.length ? h('div', { class: 'table-wrap', style: 'margin-top:12px' }, h('table', { class: 'list reply' }, h('thead', null, h('tr', null, ['줄', '구분', '협력사', '개선대책', '바뀐 내용'].map(function (x) { return h('th', null, x); }))), tb))
        : h('p', null, '바뀐 칸이 없습니다. 이미 반영했거나 협력사가 고친 것이 없습니다.'),
      d.notInFile.length ? h('p', { class: 'note' }, '이 도구에는 있는데 파일에 없는 활동 ' + d.notInFile.length + '건 — 지우지 않고 그대로 둡니다.') : null,
      c.invalid ? h('p', { class: 'note' }, '「확인 필요」 줄은 반영하지 않습니다. 협력사에 고쳐 달라고 하거나 엑셀에서 고친 뒤 다시 올리세요.') : null,
      h('div', { class: 'form-grid', style: 'margin-top:12px' }, field('반영한 사람', by, '이력에 남습니다(선택)')),
      h('div', { class: 'btn-row', style: 'margin-top:10px' },
        h('button', { type: 'button', class: 'btn btn-primary', id: 'replyApply', disabled: !applyN && !c.conflicts, onclick: applyReply }, applyN ? '반영 (' + c.changed + '행 · 새 활동 ' + c.new + '건)' : '반영할 것 없음'),
        h('button', { type: 'button', class: 'btn', onclick: function () { ui.reply = { file: null, parsed: null, diff: null, take: {}, error: '' }; render(); } }, '그만두기')));
  }
  function applyReply() {
    var st = ui.reply, at = nowIso();
    db.replyBy = (document.querySelector('#replyPreview input[aria-label="반영한 사람"]') || {}).value || db.replyBy || '';
    var res = L.applyProgress(db.plans, st.diff, { take: st.take, at: at, today: today() });
    var c = st.diff.counts;
    db.updates = (db.updates || []).concat([{ at: at, file: st.file.name, hash: st.file.hash, by: db.replyBy.trim(), suppliers: st.diff.suppliers, changed: res.changed, fields: res.fields, added: res.added,
      conflicts: c.conflicts, invalid: c.invalid, unknown: c.unknown, skipped: res.skipped, details: res.details.slice(0, 200) }]).slice(-200);
    save();
    ui.reply = { file: null, parsed: null, diff: null, take: {}, error: '' };
    render();
    toast('반영했습니다 — 바뀐 행 ' + res.changed + '건(' + res.fields + '칸) · 새 활동 ' + res.added + '건' + (res.skipped ? ' · 건너뜀 ' + res.skipped + '건' : '') + '.');
  }
  function replyHistory() {
    var list = (db.updates || []).slice().reverse();
    var card = h('section', { class: 'card', id: 'replyHistory' }, h('h2', null, '반영 이력 ' + list.length + '건'));
    if (!list.length) { card.appendChild(h('p', { class: 'note' }, '아직 반영한 회신 파일이 없습니다.')); return card; }
    card.appendChild(h('div', { class: 'table-wrap' }, h('table', { class: 'list' },
      h('thead', null, h('tr', null, ['반영 일시', '파일', '협력사', '반영한 사람', '바뀐 행', '바뀐 칸', '새 활동', '충돌', '확인 필요', '내용'].map(function (x) { return h('th', null, x); }))),
      h('tbody', null, list.map(function (u) {
        return h('tr', null, h('td', { class: 'nw' }, u.at.replace('T', ' ').slice(0, 16)), h('td', null, u.file), h('td', null, (u.suppliers || []).join(', ')), h('td', null, u.by || ''),
          h('td', { class: 'num' }, u.changed), h('td', { class: 'num' }, u.fields), h('td', { class: 'num' }, u.added), h('td', { class: 'num' }, u.conflicts || 0), h('td', { class: 'num' }, u.invalid || 0),
          h('td', null, (u.details || []).length ? h('details', null, h('summary', null, (u.details || []).length + '행'), h('ul', { class: 'diff' }, u.details.map(function (d) {
            return h('li', null, h('b', null, d.supplier + ' · ' + d.action), ' — ', d.changes.map(function (c) { return c[0] + ': ' + (c[1] === '' ? '(빈칸)' : c[1]) + ' → ' + c[2]; }).join(' / '));
          }))) : ''));
      })))));
    card.appendChild(h('p', { class: 'note' }, '이력은 이 브라우저와 JSON 백업에 함께 남습니다.'));
    return card;
  }

  // ══ 설정 ═══════════════════════════════════════════════════
  function renderSettings(main) {
    main.appendChild(pageHead('설정', '메일 받는 사람·초안, 불량유형 10종의 키워드·기본 대책, 담당부서, 개선 필요 항목 기준, 사내망·HTML 파일로 쓰기, AI 연결(선택), 데이터 백업'));
    main.appendChild(singleFileCard());
    main.appendChild(mailSettingsCard());
    main.appendChild(dictCard());
    var deptBox = h('textarea', { rows: '2' }); deptBox.value = db.depts.join(', ');
    main.appendChild(h('section', { class: 'card', id: 'deptCard' }, h('h2', null, '담당부서 목록 (협력사 내부 담당부서)'),
      h('p', { class: 'note' }, '협력사 안에서 개선활동을 맡을 부서입니다. 기본 = 품질 · 생산기술 · 영업 · 설계 · 구매. 쉼표로 나누어 더하거나 고칠 수 있고, 오른쪽 칸의 담당부서 고르기 · Excel 의 고르기 목록 · AI 프롬프트에 쓰입니다.'),
      field('담당부서', deptBox),
      h('div', { class: 'btn-row', style: 'margin-top:10px' }, h('button', { type: 'button', class: 'btn btn-primary', onclick: function () {
        var d = L.parseKeywords(deptBox.value); if (!d.length) { toast('부서를 하나 이상 적어 주세요.', true); return; }
        db.depts = d; save(); toast('저장했습니다.');
      } }, '저장'), h('button', { type: 'button', class: 'btn', onclick: function () { db.depts = L.DEFAULT_DEPTS.slice(); save(); render(); toast('기본 부서(품질 · 생산기술 · 영업 · 설계 · 구매)로 되돌렸습니다.'); } }, '기본값으로'))));
    main.appendChild(rulesCard());
    var al = db.aliases || { suppliers: {}, parts: {} };
    main.appendChild(h('section', { class: 'card' }, h('h2', null, 'AI 프롬프트 가리기'),
      h('label', { class: 'check' }, h('input', { type: 'checkbox', checked: db.mask, onchange: function (e) { db.mask = e.target.checked; save(); toast(db.mask ? '협력사·품번을 별칭으로 가립니다.' : '프롬프트에 실제 이름이 들어갑니다.', !db.mask); } }), '협력사 이름과 품번을 별칭으로 가리기(기본·권장)'),
      h('p', { class: 'note' }, '대응표: 협력사 ' + Object.keys(al.suppliers).length + '곳 · 품번 ' + Object.keys(al.parts).length + '개. 같은 이름은 늘 같은 별칭이 됩니다. 대응표는 이 브라우저와 JSON 백업에만 있습니다.')));
    main.appendChild(window.AIPanel.settingsCard({ toast: toast }));
    main.appendChild(h('section', { class: 'card' }, h('h2', null, '사내망에서 쓰기'),
      h('ul', null,
        h('li', null, '이 도구는 외부 서버·인터넷을 부르지 않습니다. 엑셀 읽기·쓰기(SheetJS)도 도구 안에 들어 있습니다.'),
        h('li', null, '가장 쉬운 방법은 위 「HTML 파일로 받기」 한 파일입니다. 사내 PC 에 저장해 두고 더블클릭하면 인터넷 없이 열립니다.'),
        h('li', null, '데이터는 각 PC 의 브라우저 안에만 저장됩니다. 다른 사람과 나누거나 PC 를 바꿀 때는 「백업 · 복원」의 JSON·엑셀 파일을 사내 공유 폴더로 주고받으세요.'),
        h('li', null, 'AI 는 API 키 없이 씁니다 — MASTER PLAN 의 프롬프트를 복사해 Copilot(사내 허용)이나 임시 채팅의 ChatGPT · Claude · Gemini 에 붙여 넣고, 답을 다시 붙여 넣습니다. 아래 「AI 연결 설정」은 사내 AI 서버가 있을 때만 쓰는 선택 기능입니다.'))));
    main.appendChild(backupCard());
  }
  function dictCard() {
    var counts = {};
    all().forEach(function (r) { counts[r.type] = (counts[r.type] || 0) + 1; });
    var list = h('div', { class: 'dict' });
    var editors = [];
    db.types.forEach(function (t, i) {
      var name = h('input', { value: t.name, 'aria-label': '유형 이름' });
      var kw = h('textarea', { rows: '2', 'aria-label': '키워드' }); kw.value = t.keywords.join(', ');
      var dept = h('input', { value: t.dept, list: 'dl-depts2', 'aria-label': '기본 담당부서' });
      var acts = h('textarea', { rows: String(Math.max(3, t.actions.length + 1)), class: 'mono', 'aria-label': '기본 대책' }); acts.value = L.formatActionLines(t.actions);
      editors.push({ name: name, kw: kw, dept: dept, acts: acts });
      function move(d) { var j = i + d; if (j < 0 || j >= db.types.length) return; readEditors(); var x = db.types[i]; db.types[i] = db.types[j]; db.types[j] = x; save(); render(); }
      list.appendChild(h('details', { class: 'dict-item' },
        h('summary', null, h('span', { class: 'dict-n' }, String(i + 1)), h('strong', null, t.name), h('span', { class: 'note' }, ' · 키워드 ' + t.keywords.length + '개 · 대책 ' + t.actions.length + '개 · 분류 ' + (counts[t.name] || 0) + '건')),
        h('div', { class: 'form-grid' }, field('유형 이름', name), field('기본 담당부서', dept)),
        field('키워드 (쉼표로 나눔)', kw, '띄어쓰기·대소문자는 무시합니다. 여러 유형이 걸리면 가장 긴 키워드가 이기고, 길이가 같으면 위에 있는 유형이 이깁니다.'),
        field('기본 대책 — 한 줄에 「대책 | 담당부서 | 시작(기준일부터 몇 주 뒤) | 기간(주)」', acts),
        h('div', { class: 'btn-row' },
          h('button', { type: 'button', class: 'btn btn-sm2', disabled: i === 0, onclick: function () { move(-1); } }, '위로'),
          h('button', { type: 'button', class: 'btn btn-sm2', disabled: i === db.types.length - 1, onclick: function () { move(1); } }, '아래로'),
          h('button', { type: 'button', class: 'btn btn-sm2 btn-danger', onclick: function () {
            confirmBox('유형 지우기', '「' + t.name + '」을 사전에서 지웁니다. 이 유형으로 분류되던 ' + (counts[t.name] || 0) + '건은 다른 유형이나 미분류가 됩니다.', '지우기', function () { readEditors(); db.types.splice(i, 1); save(); render(); }, true);
          } }, '지우기'))));
    });
    function readEditors() {
      var errs = [], names = {};
      editors.forEach(function (e, i) {
        var n = e.name.value.trim();
        if (!n || n === L.UNCLASSIFIED) errs.push((i + 1) + '번 유형 이름을 적어 주세요(「미분류」는 쓸 수 없음).');
        else if (names[n]) errs.push('유형 이름 「' + n + '」이 두 번 있습니다.');
        names[n] = 1;
      });
      if (errs.length) { toast(errs[0], true); return false; }
      var renamed = {};
      editors.forEach(function (e, i) {
        var t = db.types[i], n = e.name.value.trim();
        if (t.name !== n) renamed[t.name] = n;
        t.name = n; t.keywords = L.parseKeywords(e.kw.value); t.dept = e.dept.value.trim() || '협력사'; t.actions = L.parseActionLines(e.acts.value, t.dept);
      });
      // 이름을 바꾸면 직접 고른 유형·계획 행도 따라갑니다
      if (Object.keys(renamed).length) {
        db.records.forEach(function (r) { if (r.typeManual && renamed[r.typeManual]) r.typeManual = renamed[r.typeManual]; });
        Object.keys(db.plans).forEach(function (s) { db.plans[s].rows.forEach(function (r) { if (renamed[r.type]) r.type = renamed[r.type]; }); });
      }
      return true;
    }
    return h('section', { class: 'card' }, h('h2', null, '불량유형 키워드 사전 · 기본 대책'),
      h('p', { class: 'note' }, '불량유형 10종은 부품품질팀이 정한 분류입니다(위치/치수 · 간섭/정렬 · 이종/누락 · 용접불량 · 누유/누수 · 미조임 · 성능이상 · 도장/외관불량 · 녹/발청 · 기타). 등록할 때 유형을 고르지 않았거나 붙여 넣은 표에 유형 칸이 없으면, 내용에 든 키워드로 나눕니다. 키워드와 기본 대책(Master Plan 초안에 들어감)은 가정이니 현장 표현에 맞게 고쳐 주세요.'),
      datalist('dl-depts2', db.depts),
      list,
      h('div', { class: 'btn-row', style: 'margin-top:12px' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { if (readEditors()) { save(); toast('사전을 저장했습니다. 분류를 다시 계산했습니다.'); render(); } } }, '사전 저장'),
        h('button', { type: 'button', class: 'btn', onclick: function () {
          if (!readEditors()) return;
          var n = '새 유형', k = 1; while (L.typeNames(db.types).indexOf(n) >= 0) n = '새 유형 ' + (++k);
          db.types.push({ name: n, dept: '협력사', keywords: [], actions: [] }); save(); render();
        } }, '유형 더하기'),
        h('button', { type: 'button', class: 'btn btn-danger', onclick: function () {
          confirmBox('기본 사전으로 되돌리기', '불량유형 사전(10종)과 기본 대책을 처음 값으로 되돌립니다. 직접 더한 키워드·유형은 사라집니다.', '되돌리기', function () { db.types = L.defaultTypes(); save(); render(); }, true);
        } }, '기본값으로')));
  }
  function rulesCard() {
    var r = db.rules, ins = {};
    function num(k, label, hint, min) { ins[k] = h('input', { type: 'number', min: String(min || 1), value: String(r[k]), inputmode: 'numeric' }); return field(label, ins[k], hint); }
    return h('section', { class: 'card', id: 'rulesCard' }, h('h2', null, '개선 필요 항목 기준'),
      h('p', { class: 'note' }, '불량이 1건이라도 개선 대상입니다. 그 가운데 아래 두 경우를 우선 개선 항목으로 올립니다(2026-10-27 부품품질팀 기준): ① 협력사마다 불량유형 WORST 1·2·3 ② 같은 품번 2건 이상.'),
      h('div', { class: 'form-grid' },
        num('worstN', '① 불량유형 WORST — 협력사마다 몇 위까지', '기본 3 (WORST 1·2·3)'),
        num('partRepeatMin', '② 동일품번 — 같은 품번 몇 건 이상', '기본 2', 2)),
      h('div', { class: 'btn-row', style: 'margin-top:10px' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: function () {
          var o = {}; Object.keys(ins).forEach(function (k) { o[k] = ins[k].value; });
          db.rules = L.cleanRules(o); save(); toast('기준을 저장했습니다.'); render();
        } }, '저장'),
        h('button', { type: 'button', class: 'btn', onclick: function () { db.rules = L.cleanRules({}); save(); render(); } }, '기본값으로')));
  }
  // ══ 아웃룩 메일 초안 (4차) ════════════════════════════════════
  /* 브라우저는 아웃룩을 직접 열거나 첨부를 넣을 수 없습니다. 그래서 받는 사람 · 참조 · 제목 · 본문 · MASTER PLAN 엑셀을 담은
     메일 파일(.eml)을 내려받게 하고, 그 파일을 열면 아웃룩이 고칠 수 있는 새 메일(X-Unsent)로 엽니다. 첨부 없는 mailto: 도 함께 둡니다. */
  function peopleText(list) { return list.length ? list.map(function (p) { return p.name ? p.name + ' <' + p.email + '>' : p.email; }).join(', ') : '(없음)'; }
  function openMail(s) {
    db.mail = L.cleanMail(db.mail);
    var m = db.mail, rcp = L.mailFor(m, s), plan = planOf(s) || { rows: [] };
    var name = reportName([s]);
    if (!ui.mail.due || ui.mail.due < today()) ui.mail.due = L.addDays(today(), 7);
    var due = dateInput(ui.mail.due, { 'aria-label': '회신 마감일' });
    var subj = h('input', { 'aria-label': '메일 제목' }), body = h('textarea', { rows: '9', 'aria-label': '메일 본문' });
    function fill() { var v = L.mailVars({ supplier: s, due: due.value, today: today(), rows: plan.rows, fileName: name, fromName: m.from.name }); subj.value = L.fillTemplate(m.subject, v); body.value = L.fillTemplate(m.body, v); }
    fill();
    due.addEventListener('change', function () { ui.mail.due = due.value; fill(); });
    var errs = L.checkMail(m, s);
    function goSettings() { closeDialog(); ui.mail.per = rcp.override ? s : ''; location.hash = '#/settings'; setTimeout(function () { var c = document.getElementById('mailCard'); if (c) c.scrollIntoView({ block: 'start' }); }, 60); }
    function xlsxBytes() { var book = L.reportBook(reportCtx([s])); return XO.build(XLSX, book.sheets, book.meta); }
    function makeEml() {
      var eml = L.buildEml({ from: m.from.email ? m.from : null, to: rcp.to, cc: rcp.cc, subject: subj.value, body: body.value,
        attachments: [{ name: name, type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', bytes: xlsxBytes() }] });
      download('메일_' + L.safeFileName(s) + tag() + '_' + today() + '.eml', new Blob([eml], { type: 'message/rfc822' }));
      closeDialog();
      toast('메일 파일(.eml)을 내려받았습니다. 더블클릭하면 아웃룩이 새 메일로 엽니다 — 내용을 확인하고 「보내기」를 누르세요.');
    }
    function openMailto() {
      var url = L.mailtoUrl({ to: rcp.to, cc: rcp.cc, subject: subj.value, body: body.value });
      if (url.length > 1900) toast('본문이 길어 메일 프로그램이 일부를 자를 수 있습니다. .eml 을 권합니다.', true);
      location.href = url;
    }
    openDialog('메일 만들기 — ' + s, [
      h('p', { class: 'alert info' }, '브라우저는 보안상 아웃룩을 직접 열거나 파일을 첨부할 수 없습니다. 「아웃룩 메일 파일(.eml) 받기」를 누르면 받는 사람 · 참조 · 제목 · 본문 · MASTER PLAN 엑셀이 담긴 메일 파일을 내려받습니다. 그 파일을 더블클릭하면 아웃룩이 보내기 전 새 메일로 엽니다.'),
      errs.length ? h('div', null, errs.map(function (e) { return h('p', { class: 'alert warn' }, e); })) : null,
      h('dl', { class: 'mail-rcp' }, h('dt', null, '받는 사람'), h('dd', null, peopleText(rcp.to)), h('dt', null, '참조'), h('dd', null, peopleText(rcp.cc)),
        h('dt', null, '첨부'), h('dd', null, name + (plan.rows.length ? ' (MASTER PLAN ' + plan.rows.length + '건)' : ' — MASTER PLAN 이 아직 없어 불량목록만 들어갑니다'))),
      h('p', { class: 'note' }, rcp.override ? s + ' 전용 받는 사람을 씁니다. ' : '설정의 기본 받는 사람을 씁니다. ', h('a', { href: '#/settings', onclick: function (e) { e.preventDefault(); goSettings(); } }, '받는 사람 · 초안 바꾸기')),
      field('회신 마감일 — 본문의 {마감}', due),
      field('제목', subj), field('본문 (여기서 고친 것은 이번 메일에만 들어갑니다)', body)
    ], [
      h('button', { type: 'button', class: 'btn', onclick: closeDialog }, '그만두기'),
      h('button', { type: 'button', class: 'btn', onclick: function () { downloadReport([s]); } }, '엑셀만 받기'),
      h('button', { type: 'button', class: 'btn', disabled: !rcp.to.length, onclick: openMailto, title: '첨부 없이 기본 메일 프로그램을 엽니다' }, '메일 프로그램 열기(첨부 없음)'),
      h('button', { type: 'button', class: 'btn btn-primary', id: 'emlBtn', disabled: !!errs.length, onclick: makeEml }, icon('down'), '아웃룩 메일 파일(.eml) 받기')
    ]);
  }
  function peopleEditor(list, label) {
    var rows = [];
    var el = h('div', { class: 'people' }, list.map(function (p, i) {
      var n = h('input', { value: p.name, placeholder: '이름', 'aria-label': label + (i + 1) + ' 이름', autocomplete: 'off' });
      var e = h('input', { type: 'email', value: p.email, placeholder: 'name@company.com', 'aria-label': label + (i + 1) + ' EMAIL', autocomplete: 'off' });
      rows.push([n, e]);
      return h('div', { class: 'people-row' }, h('span', { class: 'people-k' }, label + (i + 1)), n, e);
    }));
    return { el: el, read: function () { return rows.map(function (r) { return { name: r[0].value.trim(), email: r[1].value.trim() }; }); } };
  }
  function mailSettingsCard() {
    db.mail = L.cleanMail(db.mail);
    var m = db.mail;
    var fromN = h('input', { value: m.from.name, placeholder: '예: 부품품질팀 이준호', 'aria-label': '보내는 사람 이름' });
    var fromE = h('input', { type: 'email', value: m.from.email, placeholder: '(비우면 아웃룩 기본 계정)', 'aria-label': '보내는 사람 EMAIL' });
    var to = peopleEditor(m.to, '수신'), cc = peopleEditor(m.cc, '참조');
    var subj = h('input', { value: m.subject, 'aria-label': '메일 제목 초안' });
    var body = h('textarea', { rows: '9', 'aria-label': '메일 본문 초안' }); body.value = m.body;
    function readAll() {
      var o = { from: { name: fromN.value.trim(), email: fromE.value.trim() }, to: to.read(), cc: cc.read(), subject: subj.value, body: body.value, per: m.per };
      var bad = o.to.concat(o.cc, [o.from]).filter(function (p) { return p.email && !L.isEmail(p.email); });
      if (bad.length) { toast('EMAIL 모양이 아닙니다: ' + bad[0].email, true); return null; }
      return L.cleanMail(o);
    }
    // 협력사별 받는 사람(선택)
    var sups = L.suppliersOf(db.records.concat(Object.keys(m.per).map(function (s) { return { supplier: s }; })));
    if (ui.mail.per && sups.indexOf(ui.mail.per) < 0) ui.mail.per = '';
    var psel = select([['', '협력사 고르기']].concat(sups.map(function (s) { return [s, s + (m.per[s] ? ' (따로 정함)' : '')]; })), ui.mail.per, { 'aria-label': '협력사별 받는 사람' });
    psel.addEventListener('change', function () { ui.mail.per = psel.value; render(); setTimeout(function () { var c = document.getElementById('mailPer'); if (c) c.scrollIntoView({ block: 'nearest' }); }, 30); });
    var perBox = h('div', { id: 'mailPer' }, field('협력사', psel, '고른 협력사에만 다른 수신 · 참조를 씁니다. 수신을 비우면 기본 수신, 참조를 비우면 기본 참조를 씁니다.'));
    if (ui.mail.per) {
      var cur = m.per[ui.mail.per] || { to: L.defaultMail().to, cc: L.defaultMail().cc };
      var pto = peopleEditor(cur.to, '수신'), pcc = peopleEditor(cur.cc, '참조');
      add(perBox, [h('div', { class: 'grid-2' }, pto.el, pcc.el), h('div', { class: 'btn-row' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: function () {
          var x = { to: pto.read(), cc: pcc.read() };
          var bad = x.to.concat(x.cc).filter(function (p) { return p.email && !L.isEmail(p.email); });
          if (bad.length) { toast('EMAIL 모양이 아닙니다: ' + bad[0].email, true); return; }
          var per = Object.assign({}, m.per); per[ui.mail.per] = x;
          db.mail = L.cleanMail(Object.assign({}, m, { per: per })); save(); render(); toast(ui.mail.per + ' 받는 사람을 저장했습니다.');
        } }, ui.mail.per + ' 받는 사람 저장'),
        m.per[ui.mail.per] ? h('button', { type: 'button', class: 'btn btn-danger', onclick: function () {
          var per = Object.assign({}, m.per); delete per[ui.mail.per]; db.mail = L.cleanMail(Object.assign({}, m, { per: per })); save(); render(); toast('기본 받는 사람을 쓰도록 되돌렸습니다.');
        } }, '따로 정한 것 지우기') : null)]);
    }
    var perList = Object.keys(m.per);
    return h('section', { class: 'card', id: 'mailCard' }, h('h2', null, '메일 받는 사람 · 초안 (아웃룩)'),
      h('p', { class: 'note' }, 'MASTER PLAN 이나 보고서 화면에서 협력사를 골라 「메일 만들기」를 누르면, 여기 적은 받는 사람 · 참조 · 제목 · 본문으로 아웃룩 메일 파일(.eml)을 만들고 그 협력사의 MASTER PLAN 엑셀을 첨부합니다. 브라우저는 아웃룩을 직접 움직일 수 없어 파일로 건넵니다. 주소는 이 브라우저와 JSON 백업에만 저장됩니다.'),
      h('h3', null, '보내는 사람 (선택)'), h('div', { class: 'form-grid' }, field('이름 — 본문의 {보내는사람}', fromN), field('EMAIL', fromE, '비워 두면 아웃룩이 기본 계정으로 채웁니다')),
      h('h3', null, '기본 받는 사람'), h('div', { class: 'grid-2' }, to.el, cc.el),
      h('h3', null, '메일 초안'), field('제목', subj), field('본문', body),
      h('p', { class: 'note' }, '자리표시: ' + L.MAIL_VARS.map(function (v) { return '{' + v[0] + '} = ' + v[1]; }).join(' · ')),
      h('div', { class: 'btn-row' },
        h('button', { type: 'button', class: 'btn btn-primary', id: 'mailSave', onclick: function () { var o = readAll(); if (!o) return; db.mail = o; save(); toast('메일 설정을 저장했습니다.'); render(); } }, '저장'),
        h('button', { type: 'button', class: 'btn', onclick: function () { subj.value = L.DEFAULT_MAIL.subject; body.value = L.DEFAULT_MAIL.body; toast('제목 · 본문을 기본 초안으로 바꿨습니다. 「저장」을 눌러야 남습니다.'); } }, '제목 · 본문 기본값으로')),
      h('h3', { style: 'margin-top:18px' }, '협력사별 받는 사람 (선택)' + (perList.length ? ' — ' + perList.length + '곳 따로 정함: ' + perList.join(', ') : '')),
      perBox);
  }
  /* HTML 한 파일 — 사내 보안망에서 GitHub 에 들어가기 어려워(2026-10-27) 사내 PC 에 저장해 두고 더블클릭해 씁니다 */
  function singleFileCard() {
    var steps = h('ol', { class: 'steps' },
      h('li', null, '아래 단추로 ', h('code', null, SINGLE_NAME), ' 한 파일을 받아 사내 PC 의 알맞은 폴더(예: 문서\\품질)에 둡니다. 인터넷이 되는 PC 에서 받아 사내 메일·USB 로 옮겨도 됩니다.'),
      h('li', null, '파일을 더블클릭하면 브라우저(Edge·Chrome)로 열립니다. 설치·인터넷 연결이 필요 없고, 외부 주소를 부르지 않습니다.'),
      h('li', null, '입력한 데이터는 그 PC 의 그 브라우저 안에만 저장됩니다. 다른 브라우저로 열면 빈 화면에서 시작합니다. PC 를 바꿀 때는 「백업 · 복원」의 JSON 을 받아 새 PC 에서 되살립니다.'),
      h('li', null, '새 판으로 바꿀 때: 새 파일을 받아 옛 파일 자리에 덮어쓰고(같은 이름 · 같은 폴더) 다시 엽니다. 같은 브라우저로 열면 쌓아 둔 데이터가 그대로 이어집니다. 바꾸기 전에 JSON 백업을 한 번 받아 두세요.'));
    return h('section', { class: 'card', id: 'singleCard' }, h('h2', null, 'HTML 파일 한 개로 쓰기 (사내 PC)'),
      SINGLE ? h('p', { class: 'alert info' }, '지금 열린 것이 HTML 한 파일 판입니다(판 ' + APP_VERSION + '). 새 판은 인터넷이 되는 PC 에서 이 도구 사이트(처음 이 파일을 받은 주소)를 열어 「설정 → HTML 파일로 받기」로 받습니다.')
        : h('div', { class: 'btn-row' }, h('a', { class: 'btn btn-primary', id: 'singleDownload', href: 'dist/' + SINGLE_NAME, download: SINGLE_NAME }, icon('down'), 'HTML 파일로 받기 (판 ' + APP_VERSION + ')')),
      steps);
  }
  render();
})();
