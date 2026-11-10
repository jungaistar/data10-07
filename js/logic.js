/* 협력사 공정불량 분석 · 품질개선 Master Plan — 순수 로직 (화면·저장소와 무관, node 에서도 불러 씀)

   1. 날짜·주차          5. 개선 필요 항목(WORST·동일품번)   9. AI 프롬프트·답 읽기(별칭 가리기)
   2. 엑셀 열 찾기·읽기   6. Master Plan 초안                  10. Excel 추출(불량목록·MASTER PLAN) · 10-2. 진행사항 업로드
   3. 불량유형 분류       7. 진행상태·간트 칸·추진 집계        11. 저장 형식
   4. 분석(건수·유형·월별·반복 품번)   8. 계획 확정
*/
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.SQLogic = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';
  /* 판(版) — index.html 의 ?v= · app.js 의 APP_VERSION 과 같아야 합니다(scripts/bump-version.mjs 로 함께 올림).
     브라우저가 옛 logic.js 와 새 app.js 를 섞어 쓰면 app.js 가 이 값을 보고 계산을 멈춥니다. */
  var VERSION = '2026-11-10.7';

  function str(v) { return v == null ? '' : String(v).trim(); }
  function uniq(a) { var s = {}, o = []; a.forEach(function (x) { if (!s[x]) { s[x] = 1; o.push(x); } }); return o; }
  function round(n, d) { var p = Math.pow(10, d || 0); return Math.round(n * p) / p; }
  function clone(o) { return JSON.parse(JSON.stringify(o)); }

  /* ── 1. 날짜·주차 ─────────────────────────────────────────── */
  function pad(n) { return (n < 10 ? '0' : '') + n; }
  function ymd(y, m, d) { return y + '-' + pad(m) + '-' + pad(d); }
  function validYmd(y, m, d) {
    if (!(y >= 1990 && y <= 2100 && m >= 1 && m <= 12 && d >= 1 && d <= 31)) return '';
    var t = new Date(Date.UTC(y, m - 1, d));
    return t.getUTCMonth() === m - 1 ? ymd(y, m, d) : '';
  }
  /* 엑셀 날짜 일련번호 · Date · 「2026-10-05」「2026.10.5」「2026/10/05」「2026년 10월 5일」「26.10.05」「20261005」 */
  function parseDate(v) {
    if (v == null || v === '') return '';
    if (v instanceof Date) return isNaN(v) ? '' : validYmd(v.getFullYear(), v.getMonth() + 1, v.getDate());
    if (typeof v === 'number') {
      if (v > 20000 && v < 80000) { var t = new Date(Date.UTC(1899, 11, 30) + Math.floor(v) * 86400000); return ymd(t.getUTCFullYear(), t.getUTCMonth() + 1, t.getUTCDate()); }
      if (v >= 19900101 && v <= 21001231) return parseDate(String(v));
      return '';
    }
    var s = str(v), m;
    if ((m = /^(\d{4})\s*[-./년]\s*(\d{1,2})\s*[-./월]\s*(\d{1,2})/.exec(s))) return validYmd(+m[1], +m[2], +m[3]);
    if ((m = /^(\d{2})[-./](\d{1,2})[-./](\d{1,2})$/.exec(s))) return validYmd(2000 + +m[1], +m[2], +m[3]);
    if ((m = /^(\d{4})(\d{2})(\d{2})$/.exec(s))) return validYmd(+m[1], +m[2], +m[3]);
    return '';
  }
  function toDateStr(d) { return ymd(d.getFullYear(), d.getMonth() + 1, d.getDate()); }
  function utc(s) { return Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10)); }
  function fromUtc(t) { var d = new Date(t); return ymd(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()); }
  function addDays(s, n) { return fromUtc(utc(s) + n * 86400000); }
  function daysBetween(a, b) { return Math.round((utc(b) - utc(a)) / 86400000); }
  function dow(s) { return new Date(utc(s)).getUTCDay(); } // 0=일
  function mondayOf(s) { var w = dow(s); return addDays(s, w === 0 ? -6 : 1 - w); }
  function monthOf(s) { return s ? s.slice(0, 7) : ''; }
  function monthAdd(ym, k) { var y = +ym.slice(0, 4), m = +ym.slice(5, 7) - 1 + k; y += Math.floor(m / 12); m = ((m % 12) + 12) % 12; return y + '-' + pad(m + 1); }
  function monthLast(ym) { return addDays(monthAdd(ym, 1) + '-01', -1); }
  /* 주차 — 월~일 한 주는 「목요일이 든 달」의 주로 셉니다(2026-10-26 ~ 11-01 → 2026년 10월 5주차) */
  function weekInfo(s) {
    var mon = mondayOf(s), thu = addDays(mon, 3);
    var y = +thu.slice(0, 4), m = +thu.slice(5, 7), w = Math.floor((+thu.slice(8, 10) - 1) / 7) + 1;
    return { monday: mon, sunday: addDays(mon, 6), year: y, month: m, week: w, label: y + '년 ' + m + '월 ' + w + '주차' };
  }

  /* ── 2. 엑셀 열 찾기·읽기 ──────────────────────────────────── */
  /* aliases 는 앞에 있을수록 먼저 고릅니다(같은 표에 「협력회사」(코드)와 「협력회사명」(이름)이 함께 있으면 이름 쪽).
     4차(2026-10-27): 울산캠퍼스(문제점 번호 · 발생 일자 · 협력회사명 · 문제점 상세 · 조치방안 상세 …) ·
     인천캠퍼스(QM noti No. · Issued Date · Vendor Name · Material · Problem Description · Corrective Action …) 두 시스템의 열 이름을 더했습니다. */
  var FIELDS = [
    { id: 'date', label: '발생일자', req: true, aliases: ['발생일자', '발생일', '발생날짜', '일자', '날짜', '접수일', '접수일자', '발견일', '발견일자', 'issueddate', 'issuedate', 'date'] },
    { id: 'partName', label: '품명', req: false, aliases: ['품명', '품목명', '부품명', '품목', '부품상세', 'materialdescription', 'partname'] },
    { id: 'partNo', label: '품번', req: false, aliases: ['품번', '부품번호', '품목번호', '품목코드', 'material', 'partno', 'p/n', 'pn'] },
    { id: 'problem', label: '내용(문제점)', req: true, aliases: ['문제점', '문제점상세', '불량내용', '불량현상', '문제내용', '현상', '내용', '불량명', 'problemdescription', 'defect'] },
    { id: 'supplier', label: '협력사', req: true, aliases: ['협력사', '협력사명', '협력회사명', '협력업체', '업체', '업체명', '공급사', '거래처', 'vendorname', 'suppliername', 'supplier', '협력회사', 'vendor'] },
    { id: 'type', label: '불량유형(선택)', req: false, aliases: ['불량유형', '유형', '불량구분', '불량분류', '불량코드', 'type', 'defecttype'] },
    { id: 'qty', label: '수량(선택)', req: false, aliases: ['수량', '불량수량', 'qty'] },
    { id: 'action', label: '조치(선택)', req: false, aliases: ['조치', '조치내용', '조치사항', '조치방안상세', '조치방안', '처리', '처리내용', '처리결과', 'correctiveaction', 'action'] },
    { id: 'srcNo', label: '관리번호(선택)', req: false, aliases: ['관리번호', '문제점번호', '접수번호', 'qmnotino', 'qmnotificationno', 'notificationno', 'notino'] },
    { id: 'campus', label: '캠퍼스(선택)', req: false, aliases: ['캠퍼스', '사업장', 'campus'] }
  ];
  /* 캠퍼스 — 회사는 통합됐지만 시스템은 통합 전이라 울산 · 인천이 RAW DATA 를 따로 뽑습니다(2026-10-27 수강생 요청) */
  var CAMPUSES = ['울산', '인천'];
  function campusLabel(c) { return c ? c + '캠퍼스' : '구분 없음'; }
  function normCampus(v) { var s = str(v); if (!s) return ''; for (var i = 0; i < CAMPUSES.length; i++) if (s.indexOf(CAMPUSES[i]) >= 0) return CAMPUSES[i]; return s.replace(/캠퍼스$/, ''); }
  /* 엑셀 붙여넣기에 머리행이 없을 때의 열 순서 = 불량 등록 화면·붙여넣기 양식의 순서 */
  var PASTE_ORDER = ['date', 'supplier', 'partNo', 'partName', 'type', 'problem', 'qty', 'action'];
  function fieldLabel(id) { var f = FIELDS.filter(function (x) { return x.id === id; })[0]; return f ? f.label.replace(/\(.*\)$/, '') : id; }
  function normHead(v) { return str(v).replace(/\(.*?\)|\[.*?\]/g, '').replace(/[\s·.:_\-]/g, '').toLowerCase(); }
  /* 한 줄의 머리 칸 → { 항목id: 열번호 }. ① 이름이 꼭 같은 칸(별칭 순서가 앞인 것부터) ② 이름이 별칭으로 시작하는 칸.
     줄 전체를 ① 로 먼저 훑으므로, 왼쪽의 「Vendor Code」「문제점 번호」「조치 담당자명」이 오른쪽의 「Vendor Name」「문제점 상세」「조치방안 상세」를 가로채지 않습니다. */
  function matchRow(row) {
    var heads = (row || []).map(normHead), map = {}, used = {};
    FIELDS.forEach(function (f) {
      var best = -1, bestRank = 1e9;
      heads.forEach(function (n, c) { if (!n || used[c]) return; var k = f.aliases.indexOf(n); if (k >= 0 && k < bestRank) { bestRank = k; best = c; } });
      if (best >= 0) { map[f.id] = best; used[best] = 1; }
    });
    FIELDS.forEach(function (f) {
      if (map[f.id] != null) return;
      for (var c = 0; c < heads.length; c++) {
        var n = heads[c];
        if (n && !used[c] && f.aliases.some(function (a) { return a.length >= 2 && n.indexOf(a) === 0; })) { map[f.id] = c; used[c] = 1; return; }
      }
    });
    return map;
  }
  /* 위 15행 안에서 머리행을 찾습니다(제목 줄이 있는 대장 형식도 읽기 위해) */
  function detectColumns(aoa) {
    var best = { headerRow: -1, map: {}, score: 0 };
    for (var r = 0; r < Math.min(15, aoa.length); r++) {
      var map = matchRow(aoa[r]), score = Object.keys(map).length;
      if (score > best.score) best = { headerRow: r, map: map, score: score };
    }
    if (best.score < 2) best = { headerRow: -1, map: {}, score: 0 };
    var headers = best.headerRow >= 0 ? (aoa[best.headerRow] || []).map(str) : [];
    var missing = FIELDS.filter(function (f) { return f.req && best.map[f.id] == null; }).map(function (f) { return f.id; });
    return { headerRow: best.headerRow, map: best.map, headers: headers, missing: missing };
  }
  function parseRows(aoa, det) {
    var out = [], badDate = 0, skipped = 0;
    if (det.headerRow < 0 || det.missing.length) return { records: out, badDate: 0, skipped: 0 };
    var M = det.map;
    function cell(row, id) { return M[id] == null ? null : row[M[id]]; }
    for (var r = det.headerRow + 1; r < aoa.length; r++) {
      var row = aoa[r] || [];
      var rec = {
        date: parseDate(cell(row, 'date')),
        partName: str(cell(row, 'partName')),
        partNo: str(cell(row, 'partNo')),
        problem: str(cell(row, 'problem')).replace(/\s*\n\s*/g, ' '),
        supplier: str(cell(row, 'supplier'))
      };
      var q = Number(cell(row, 'qty'));
      if (M.qty != null && isFinite(q) && str(cell(row, 'qty')) !== '') rec.qty = q;
      if (str(cell(row, 'type'))) rec.typeIn = str(cell(row, 'type'));
      if (str(cell(row, 'action'))) rec.action = str(cell(row, 'action')).replace(/\s*\n\s*/g, ' ');
      if (str(cell(row, 'srcNo'))) rec.srcNo = str(cell(row, 'srcNo'));
      if (normCampus(cell(row, 'campus'))) rec.campus = normCampus(cell(row, 'campus'));
      if (!rec.problem && !rec.supplier && !rec.partNo) { if (row.some(function (v) { return str(v); })) skipped++; continue; }
      // 내용이 비었어도 불량유형 칸이 있으면 그것을 내용으로 씁니다(유형만 적는 대장도 있어서)
      if (!rec.problem && rec.typeIn) rec.problem = rec.typeIn;
      if (!rec.problem || !rec.supplier) { skipped++; continue; }
      if (!rec.date) badDate++;
      out.push(rec);
    }
    return { records: out, badDate: badDate, skipped: skipped };
  }
  /* overrides: { 항목id: 열번호(-1 = 연결 끊기) } — 사용자가 열 연결을 고친 경우 */
  function parseSheet(aoa, opts) {
    var det = detectColumns(aoa);
    var ov = opts && opts.overrides;
    if (ov) {
      if (det.headerRow < 0 && opts.headerRow != null) { det.headerRow = opts.headerRow; det.headers = (aoa[opts.headerRow] || []).map(str); }
      Object.keys(ov).forEach(function (k) { if (ov[k] == null || ov[k] < 0) delete det.map[k]; else det.map[k] = ov[k]; });
      det.missing = FIELDS.filter(function (f) { return f.req && det.map[f.id] == null; }).map(function (f) { return f.id; });
    }
    var p = parseRows(aoa, det);
    return { det: det, records: p.records, badDate: p.badDate, skipped: p.skipped };
  }
  /* 여러 시트 중 머리행이 가장 잘 맞는 시트 */
  function pickSheet(sheets) {
    var best = null, bestScore = -1;
    sheets.forEach(function (s) {
      var d = detectColumns(s.aoa), sc = d.headerRow < 0 ? -1 : Object.keys(d.map).length * 10 - d.missing.length * 20;
      if (sc > bestScore) { bestScore = sc; best = s; }
    });
    return best;
  }
  /* 같은 행 = 같은 캠퍼스 · 같은 관리번호 · 발생일자 · 협력사 · 품번 · 문제점. 캠퍼스가 다르면 같은 내용이어도 다른 행입니다(중복 제거는 캠퍼스 안에서만) */
  function recordKey(r) { return [str(r.campus), str(r.srcNo).replace(/\s+/g, '').toUpperCase(), r.date, str(r.supplier).replace(/\s+/g, ''), str(r.partNo).replace(/\s+/g, '').toUpperCase(), str(r.problem).replace(/\s+/g, '')].join('|'); }
  /* 같은 행(위 recordKey)은 한 번만. 같은 날 같은 불량이 정말 두 건이면 수량 열로 구분해 주세요 */
  function mergeRecords(old, add, seq) {
    var seen = {}, out = old.slice(), added = 0, dup = 0;
    seq = seq || 0;
    old.forEach(function (r) { seen[recordKey(r)] = 1; });
    add.forEach(function (r) {
      var k = recordKey(r);
      if (seen[k]) { dup++; return; }
      seen[k] = 1;
      var c = clone(r); c.id = 'd' + (++seq); out.push(c); added++;
    });
    return { records: out, added: added, dup: dup, seq: seq };
  }
  /* 캠퍼스 RAW DATA 올리기 — 그 캠퍼스에서 올린 RAW 행(src: 'raw')만 새 파일로 바꾸고, 다른 캠퍼스와 직접 등록한 행은 그대로 둡니다.
     같은 행(recordKey)이 예전에도 있었으면 그 행에서 고른 불량유형(typeManual)을 이어 받습니다.
     → { records, added, dup, removed, carried, seq } */
  function replaceCampus(old, campus, add, seq) {
    var keep = [], gone = {}, removed = 0, carried = 0;
    (old || []).forEach(function (r) {
      if (r.campus === campus && r.src === 'raw') { removed++; if (r.typeManual) gone[recordKey(r)] = r.typeManual; }
      else keep.push(r);
    });
    var tagged = (add || []).map(function (r) { var c = clone(r); c.campus = campus; c.src = 'raw'; var k = recordKey(c); if (gone[k] && !c.typeManual) { c.typeManual = gone[k]; carried++; } return c; });
    var m = mergeRecords(keep, tagged, seq);
    return { records: m.records, added: m.added, dup: m.dup, removed: removed, carried: carried, seq: m.seq };
  }
  /* 캠퍼스별 건수 — { 울산: n, 인천: m, '': 직접 등록 등 구분 없음 } */
  function campusCounts(records) {
    var o = {}; CAMPUSES.forEach(function (c) { o[c] = 0; }); o[''] = 0;
    (records || []).forEach(function (r) { var c = str(r.campus); o[c] = (o[c] || 0) + 1; });
    return o;
  }
  function campusMatch(r, campus) { return !campus || (campus === '_none' ? !str(r.campus) : r.campus === campus); }
  /* 엑셀에서 복사한 표(탭으로 나뉜 글) → 배열의 배열. 칸 안 줄바꿈은 엑셀이 "…" 로 감싸 줍니다 */
  function splitTsv(text) {
    var s = String(text == null ? '' : text).replace(/\r\n?/g, '\n'), rows = [], row = [], cell = '', i = 0, q = false;
    while (i < s.length) {
      var ch = s.charAt(i);
      if (q) {
        if (ch === '"' && s.charAt(i + 1) === '"') { cell += '"'; i += 2; continue; }
        if (ch === '"') { q = false; i++; continue; }
        cell += ch; i++; continue;
      }
      if (ch === '"' && cell === '') { q = true; i++; continue; }
      if (ch === '\t') { row.push(cell); cell = ''; i++; continue; }
      if (ch === '\n') { row.push(cell); rows.push(row); row = []; cell = ''; i++; continue; }
      cell += ch; i++;
    }
    if (cell !== '' || row.length) { row.push(cell); rows.push(row); }
    return rows.filter(function (r) { return r.some(function (v) { return str(v); }); });
  }
  /* 붙여넣기 읽기 — 머리행(열 이름)이 있으면 그 이름으로, 없으면 PASTE_ORDER 순서로 읽습니다.
     날짜처럼 읽히는 칸이 있는 줄은 머리행으로 보지 않습니다(데이터 줄을 머리행으로 오해하지 않게) */
  function parsePaste(text) {
    var aoa = splitTsv(text);
    if (!aoa.length) return { records: [], badDate: 0, skipped: 0, header: false, rows: 0, missing: [] };
    var det = detectColumns(aoa);
    var header = det.headerRow >= 0 && !(aoa[det.headerRow] || []).some(function (v) { return parseDate(v); });
    var p;
    if (header) p = parseSheet(aoa);
    else {
      var map = {}; PASTE_ORDER.forEach(function (id, c) { map[id] = c; });
      p = parseRows([PASTE_ORDER.slice()].concat(aoa), { headerRow: 0, map: map, missing: [] });
      p.det = { headerRow: -1, map: map, missing: [] };
    }
    return { records: p.records, badDate: p.badDate, skipped: p.skipped, header: header, rows: aoa.length - (header ? det.headerRow + 1 : 0), missing: p.det.missing || [] };
  }
  /* 화면에서 직접 등록하는 한 건 — 점검 후 저장용 레코드로 */
  function checkRecordInput(o, types) {
    var e = [];
    if (!str(o.date)) e.push('발생일자를 골라 주세요.');
    else if (!parseDate(o.date)) e.push('발생일자를 읽지 못했습니다(예: 2026-10-05).');
    if (!str(o.supplier)) e.push('협력사를 적어 주세요.');
    if (!str(o.problem)) e.push('내용(문제점)을 적어 주세요.');
    if (str(o.qty) !== '' && !(isFinite(Number(o.qty)) && Number(o.qty) >= 0)) e.push('수량은 0 이상의 숫자로 적어 주세요.');
    if (str(o.type) && str(o.type) !== UNCLASSIFIED && typeNames(types).indexOf(str(o.type)) < 0) e.push('불량유형을 목록에서 골라 주세요.');
    return e;
  }
  function recordFromInput(o) {
    var r = { date: parseDate(o.date), partName: str(o.partName), partNo: str(o.partNo), problem: str(o.problem).replace(/\s*\n\s*/g, ' '), supplier: str(o.supplier) };
    if (str(o.qty) !== '') r.qty = Number(o.qty);
    if (str(o.type)) r.typeManual = str(o.type);
    if (str(o.action)) r.action = str(o.action);
    if (normCampus(o.campus)) r.campus = normCampus(o.campus);
    return r;
  }

  /* ── 3. 불량유형 분류 (키워드 사전, 설정에서 고침) ─────────────── */
  var UNCLASSIFIED = '미분류';
  /* 담당부서 = 협력사 「안」의 담당부서(2026-10-27 수강생 답: 협력사 내부 담당부서 지정 목적). 설정에서 고칩니다 */
  var DEFAULT_DEPTS = ['품질', '생산기술', '영업', '설계', '구매'];
  /* actions: [대책, 담당부서, 시작(기준일부터 몇 주 뒤), 기간(주)] — 기본 대책(가정). 현장 표현에 맞게 설정에서 고칩니다 */
  var DEFAULT_TYPES = [
    /* 수강생이 정한 불량유형 8종(2026-10-21) + 도장/외관불량 · 녹/발청(2026-10-27) = 10종. 키워드·기본 대책은 가정 — 설정에서 고칩니다. */
    { name: '위치/치수', dept: '품질', keywords: ['치수', '공차', '홀 위치', '홀위치', '위치 불량', '위치 틀림', '위치 상이', '피치', '길이', '외경', '내경', '두께', '높이', '편심', '평면도', '동심', '진원', '규격 초과', '규격 미달', '측정값', '편차', '가공', '탭', '나사산', 'burr', '바리'],
      actions: [['치수·위치 불량 부위 측정, 공정능력(Cpk) 확인', '품질', 0, 2], ['지그·기준면·공구 마모 점검, 가공 조건 재설정', '생산기술', 1, 2], ['초·중·종물 측정 기록 의무화', '품질', 2, 2], ['도면 공차·기준 위치 적정성 검토', '설계', 2, 2], ['유효성 확인(대책 후 재발 여부 추적)', '품질', 5, 4]] },
    { name: '간섭/정렬', dept: '품질', keywords: ['간섭', '정렬', '틀어짐', '어긋남', '단차', '유격', '센터', '직각', '평행', '비틀림', '닿음', '맞지 않음', '조립 불가', '조립불가', '쏠림', '오조립', '역조립', '방향 반대'],
      actions: [['간섭·정렬 불량 부위와 상대 부품 조사(조립 상태 재현)', '품질', 0, 1], ['조립 지그·위치 결정핀 점검, 정렬 기준 한도견본 제작', '생산기술', 1, 2], ['상대 부품 공차 누적 검토', '설계', 1, 2], ['조립 순서 작업표준서 개정', '생산기술', 2, 2], ['유효성 확인(대책 후 재발 여부 추적)', '품질', 4, 4]] },
    { name: '이종/누락', dept: '품질', keywords: ['이종', '오품', '혼입', '혼용', '타품', '오삽', '오출하', '품번 상이', '품번 오류', '라벨 오류', '식별 오류', '사양 상이', '누락', '결품', '빠짐', '미장착', '미부착', '미삽입', '수량 부족'],
      actions: [['이종·누락 발생 경로 조사(입고·보관·가공·조립·출하 단계별)', '품질', 0, 1], ['품번별 식별표·보관 위치 구분, 부품 키팅·체크시트 도입', '생산기술', 1, 2], ['구매 부품(2차 협력사) 식별·입고 검사 기준 확인', '구매', 1, 2], ['출하 전 라벨·바코드 대조, 누락 확인 항목 추가', '품질', 2, 2], ['유효성 확인(대책 후 재발 여부 추적)', '품질', 4, 4]] },
    { name: '용접불량', dept: '품질', keywords: ['용접', '균열', '크랙', 'crack', '갈라짐', '기공', '블로우홀', '언더컷', '스패터', '미용접', '비드', '각장', '용입', '오버랩'],
      actions: [['용접 결함 유형별 발생 위치·원인 분석(용접 조건·모재·예열)', '품질', 0, 2], ['용접 조건(전류·전압·속도·예열) 표준 재설정, 작업표준서 개정', '생산기술', 2, 2], ['용접부 외관·비파괴검사(PT/MT) 샘플링 강화', '품질', 2, 4], ['용접사 교육·자격 점검, 지그·용접 순서 점검', '생산기술', 3, 1], ['유효성 확인(대책 후 재발 여부 추적)', '품질', 6, 4]] },
    { name: '누유/누수', dept: '품질', keywords: ['누유', '누수', '리크', 'leak', '오일 샘', '물 샘', '기밀', '새어', '맺힘'],
      actions: [['기밀 시험 조건(압력·시간) 점검', '품질', 0, 2], ['실링면 가공 상태·오링·피팅 체결 확인', '생산기술', 1, 2], ['누유/누수 검사 전수화 여부 검토', '품질', 2, 2], ['유효성 확인(대책 후 재발 여부 추적)', '품질', 4, 4]] },
    { name: '미조임', dept: '품질', keywords: ['미조임', '조임 불량', '조임불량', '미체결', '체결', '토크', '풀림', '헐거', '볼트', '너트'],
      actions: [['체결 토크 관리 기준·토크 공구 교정 점검', '생산기술', 0, 2], ['체결 확인 마킹(토크 마킹)·체크시트 도입', '생산기술', 1, 2], ['체결 순서 작업표준서 개정', '생산기술', 2, 2], ['유효성 확인(대책 후 재발 여부 추적)', '품질', 4, 4]] },
    { name: '성능이상', dept: '품질', keywords: ['성능', '작동', '동작', '기능', '소음', '이음', '진동', '발열', '압력 저하', '압력 부족', '출력', '응답', '고장', '불능'],
      actions: [['성능이상 재현 시험·원인 부품 분리', '품질', 0, 2], ['출하 성능검사 항목·판정 기준 점검', '품질', 1, 2], ['설계 사양·시험 조건 검토', '설계', 2, 2], ['유효성 확인(대책 후 재발 여부 추적)', '품질', 4, 4]] },
    { name: '도장/외관불량', dept: '품질', keywords: ['도장', '도막', '도막 두께', '도색', '페인트', '도장 흐름', '찍힘', '스크래치', '긁힘', '외관', '얼룩', '광택', '핀홀', '기포', '오렌지 필'],
      actions: [['도장·외관 불량 부위·유형 조사(흐름·찍힘·스크래치·도막)', '품질', 0, 2], ['도장 조건(점도·도막 두께·건조) 표준 재설정', '생산기술', 1, 2], ['취급·적재 방법 개선(보호재·전용 대차)', '생산기술', 2, 2], ['외관 한도견본 제작·출하 검사 기준 반영', '품질', 2, 2], ['유효성 확인(대책 후 재발 여부 추적)', '품질', 5, 4]] },
    { name: '녹/발청', dept: '품질', keywords: ['녹', '녹 발생', '녹발생', '녹슴', '녹 슴', '발청', '부식', '산화', '적청', '백청'],
      actions: [['녹 발생 부위·보관 조건(습도·기간) 조사', '품질', 0, 1], ['방청 처리(방청유·방청 포장) 기준 수립', '생산기술', 1, 2], ['방청유·포장재 구매 사양 확인', '구매', 1, 2], ['납품 포장·운송 사양 협의', '영업', 2, 2], ['유효성 확인(대책 후 재발 여부 추적)', '품질', 4, 4]] },
    { name: '기타', dept: '품질', keywords: ['이물', '오염', '변형', '포장', '파손'],
      actions: [['발생 내용 확인·원인 조사', '품질', 0, 2], ['재발 방지 대책 수립·작업표준 반영', '품질', 2, 2], ['유효성 확인(대책 후 재발 여부 추적)', '품질', 4, 4]] }
  ];
  function defaultTypes() {
    return DEFAULT_TYPES.map(function (t) {
      return { name: t.name, dept: t.dept, keywords: t.keywords.slice(), actions: t.actions.map(function (a) { return { text: a[0], dept: a[1], start: a[2], weeks: a[3] }; }) };
    });
  }
  function normText(s) { return str(s).toLowerCase().replace(/\s+/g, ''); }
  /* 점수 = 맞은 키워드 가운데 가장 긴 것의 글자 수(긴 키워드가 더 구체적).
     같으면 사전에서 위에 있는 유형 — 그래서 사전 순서가 곧 우선순위입니다(「이종 볼트 혼용」 → 혼용·볼트 모두 2글자 → 위에 있는 이종/누락). */
  function classify(text, types) {
    var t = normText(text), best = null;
    (types || []).forEach(function (ty, i) {
      var hits = [], score = 0;
      (ty.keywords || []).forEach(function (k) {
        var n = normText(k);
        if (n && t.indexOf(n) >= 0 && hits.indexOf(k) < 0) { hits.push(k); score = Math.max(score, n.length); }
      });
      if (score > 0 && (!best || score > best.score)) best = { type: ty.name, hits: hits, score: score, index: i };
    });
    return best || { type: UNCLASSIFIED, hits: [], score: 0, index: -1 };
  }
  function typeNames(types) { return (types || []).map(function (t) { return t.name; }); }
  /* 유형 이름 맞추기 — 띄어쓰기·「/」「·」 같은 구분 기호는 무시(「위치·치수」「위치 / 치수」 → 위치/치수, 「용접 불량」 → 용접불량) */
  function normType(s) { return str(s).toLowerCase().replace(/[\s\/·・,.\-_()\[\]]/g, ''); }
  function findTypeName(s, types) {
    var n = normType(s), names = typeNames(types), i;
    if (!n) return '';
    for (i = 0; i < names.length; i++) if (normType(names[i]) === n) return names[i];
    if (n.length >= 2) for (i = 0; i < names.length; i++) { var m = normType(names[i]); if (m.length >= 2 && (n.indexOf(m) >= 0 || m.indexOf(n) >= 0)) return names[i]; }
    return '';
  }
  /* 순서: ① 사람이 고른 유형(typeManual) ② 엑셀의 불량유형 칸(typeIn)이 사전 이름과 맞으면 그것 ③ 내용 키워드 */
  function typeOf(rec, types) {
    if (rec.typeManual && (rec.typeManual === UNCLASSIFIED || typeNames(types).indexOf(rec.typeManual) >= 0)) return { type: rec.typeManual, hits: [], manual: true };
    var col = rec.typeIn ? findTypeName(rec.typeIn, types) : '';
    if (col) return { type: col, hits: [], manual: false, column: true };
    var c = classify(rec.problem + (rec.typeIn ? ' ' + rec.typeIn : ''), types);
    return { type: c.type, hits: c.hits, manual: false };
  }
  function enrich(records, types) {
    var names = typeNames(types);
    return records.map(function (r) {
      var t = typeOf(r, types), c = {};
      Object.keys(r).forEach(function (k) { c[k] = r[k]; });
      c.type = t.type; c.hits = t.hits; c.manual = t.manual; c.column = !!t.column; c.typeIndex = names.indexOf(t.type);
      return c;
    });
  }
  /* 설정 화면의 「대책 | 부서 | 시작주 | 기간주」 줄 ↔ actions */
  function parseActionLines(text, defDept) {
    return String(text || '').split(/\r?\n/).map(function (line) {
      var p = line.split('|').map(str);
      if (!p[0]) return null;
      var s = parseInt(p[2], 10), w = parseInt(p[3], 10);
      return { text: p[0], dept: p[1] || defDept || '품질', start: isFinite(s) && s >= 0 ? s : 0, weeks: isFinite(w) && w > 0 ? w : 2 };
    }).filter(Boolean);
  }
  function formatActionLines(actions) {
    return (actions || []).map(function (a) { return a.text + ' | ' + a.dept + ' | ' + a.start + ' | ' + a.weeks; }).join('\n');
  }
  function parseKeywords(text) { return uniq(String(text || '').split(/[,\n،、]/).map(str).filter(Boolean)); }

  /* ── 4. 분석 ─────────────────────────────────────────────── */
  /* 개선 필요(우선) 기준 — 2026-10-27 수강생 답: 두 경우만. ① 불량유형별 WORST 1·2·3 ② 동일품번 2건 이상.
     (불량이 1건이라도 개선 대상이지만, 이 두 경우를 우선 개선 항목으로 올립니다) */
  var DEFAULT_RULES = { worstN: 3, partRepeatMin: 2 };
  function cleanRules(r) {
    var o = clone(DEFAULT_RULES);
    if (!r || typeof r !== 'object') return o;
    function n(k, min, max) { var v = Number(r[k]); if (r[k] !== '' && r[k] != null && isFinite(v) && v >= min && v <= max) o[k] = Math.round(v); }
    n('worstN', 1, 20); n('partRepeatMin', 2, 999);
    return o;
  }
  function inRange(r, from, to) { return (!from || (r.date && r.date >= from)) && (!to || (r.date && r.date <= to)); }
  function countBy(rows, f) {
    var m = {};
    rows.forEach(function (r) { var k = f(r); m[k] = (m[k] || 0) + 1; });
    return Object.keys(m).map(function (k) { return { key: k, count: m[k] }; }).sort(function (a, b) { return b.count - a.count || (a.key < b.key ? -1 : 1); });
  }
  function dateBounds(rows) {
    var ds = rows.map(function (r) { return r.date; }).filter(Boolean).sort();
    return { min: ds[0] || '', max: ds[ds.length - 1] || '' };
  }
  /* rows 는 enrich() 를 거친 것. opts = { supplier, from, to } */
  function analyze(rows, opts, rules) {
    opts = opts || {}; rules = cleanRules(rules);
    var sel = rows.filter(function (r) { return (!opts.supplier || r.supplier === opts.supplier) && campusMatch(r, opts.campus) && inRange(r, opts.from, opts.to); });
    var b = dateBounds(sel);
    var from = opts.from || b.min, to = opts.to || b.max;
    var total = sel.length;
    var byType = countBy(sel, function (r) { return r.type; });
    // 파레토: 미분류는 맨 뒤
    byType.sort(function (x, y) { return (x.key === UNCLASSIFIED) - (y.key === UNCLASSIFIED) || y.count - x.count || (x.key < y.key ? -1 : 1); });
    var cum = 0;
    byType = byType.map(function (x) { cum += x.count; return { type: x.key, count: x.count, pct: total ? round(x.count / total * 100, 1) : 0, cum: total ? round(cum / total * 100, 1) : 0 }; });
    var suppliers = countBy(sel, function (r) { return r.supplier; }).map(function (x) { return { name: x.key, count: x.count }; });
    // 월별 — 빈 달도 0 으로
    var months = [];
    if (from && to) {
      var mc = {};
      sel.forEach(function (r) { if (r.date) mc[monthOf(r.date)] = (mc[monthOf(r.date)] || 0) + 1; });
      for (var ym = monthOf(from), guard = 0; ym <= monthOf(to) && guard < 240; ym = monthAdd(ym, 1), guard++) months.push({ month: ym, count: mc[ym] || 0 });
    }
    // 반복 품번 — 같은 협력사·같은 품번
    var pm = {};
    sel.forEach(function (r) {
      if (!r.partNo) return;
      var k = r.supplier + '\u0000' + r.partNo;
      var g = pm[k] || (pm[k] = { supplier: r.supplier, partNo: r.partNo, partName: r.partName, count: 0, types: [], typeCount: {}, last: '' });
      g.count++; if (g.types.indexOf(r.type) < 0) g.types.push(r.type); g.typeCount[r.type] = (g.typeCount[r.type] || 0) + 1; if (r.date > g.last) g.last = r.date; if (!g.partName) g.partName = r.partName;
    });
    var repeatParts = Object.keys(pm).map(function (k) { return pm[k]; }).filter(function (g) { return g.count >= rules.partRepeatMin; })
      .sort(function (a, b) { return b.count - a.count || (a.supplier < b.supplier ? -1 : a.supplier > b.supplier ? 1 : a.partNo < b.partNo ? -1 : 1); });
    var worst = findWorst(sel, rules);
    var issues = worst.concat(repeatParts.map(function (p) {
      // 품번의 대표 유형 = 가장 많이 나온 유형(미분류 제외) — 초안 행의 불량유형으로 씁니다
      var main = p.types.filter(function (t) { return t !== UNCLASSIFIED; }).sort(function (a, b) { return p.typeCount[b] - p.typeCount[a]; })[0] || '기타';
      p.types.sort(function (a, b) { return p.typeCount[b] - p.typeCount[a]; });
      return { kind: 'part', supplier: p.supplier, partNo: p.partNo, partName: p.partName, count: p.count, types: p.types, mainType: main, last: p.last, label: '동일품번' };
    })).sort(issueOrder);
    return {
      from: from, to: to, rows: sel, total: total,
      qty: sel.reduce(function (s, r) { return s + (typeof r.qty === 'number' ? r.qty : 0); }, 0),
      hasQty: sel.some(function (r) { return typeof r.qty === 'number'; }),
      byType: byType, suppliers: suppliers, months: months, campuses: campusCounts(sel), repeatParts: repeatParts, worst: worst, issues: issues,
      unclassified: sel.filter(function (r) { return r.type === UNCLASSIFIED; }).length
    };
  }

  /* ── 5. 개선 필요 항목 — ① 협력사마다 불량유형 WORST 1·2·3 ② 동일품번 2건 이상 ─────────────
     WORST 순위: 건수가 많은 순 → 같으면 최근 발생일이 늦은 순 → 같으면 사전 순서. 미분류는 순위에서 뺍니다.
     동일품번: 같은 협력사·같은 품번이 partRepeatMin 건 이상(불량유형은 달라도 셈). */
  function findWorst(sel, rules) {
    rules = cleanRules(rules);
    var g = {}, order = {};
    sel.forEach(function (r) {
      if (r.type === UNCLASSIFIED) return;
      var k = r.supplier + '\u0000' + r.type;
      var x = g[k] || (g[k] = { kind: 'worst', supplier: r.supplier, type: r.type, count: 0, last: '', parts: {}, problems: [] });
      x.count++;
      if (r.date && r.date > x.last) x.last = r.date;
      if (r.partNo) x.parts[r.partNo] = (x.parts[r.partNo] || 0) + 1;
      if (r.problem && x.problems.indexOf(r.problem) < 0) x.problems.push(r.problem);
      if (order[r.type] == null) order[r.type] = r.typeIndex == null ? 999 : r.typeIndex;
    });
    var bySup = {};
    Object.keys(g).forEach(function (k) { var x = g[k]; (bySup[x.supplier] || (bySup[x.supplier] = [])).push(x); });
    var out = [];
    Object.keys(bySup).forEach(function (s) {
      bySup[s].sort(function (a, b) { return b.count - a.count || (a.last < b.last ? 1 : a.last > b.last ? -1 : 0) || (order[a.type] - order[b.type]) || (a.type < b.type ? -1 : 1); })
        .slice(0, rules.worstN).forEach(function (x, i) {
          x.rank = i + 1; x.label = 'WORST ' + (i + 1); x.problems = x.problems.slice(0, 5);
          x.repeatParts = Object.keys(x.parts).filter(function (p) { return x.parts[p] >= rules.partRepeatMin; }).map(function (p) { return { partNo: p, count: x.parts[p] }; });
          delete x.parts;
          out.push(x);
        });
    });
    return out;
  }
  /* 목록 순서: 건수가 많은 순 → WORST 가 동일품번보다 먼저 → 순위 → 협력사 이름 */
  function issueOrder(a, b) {
    return b.count - a.count || ((a.kind === 'worst' ? 0 : 1) - (b.kind === 'worst' ? 0 : 1)) || ((a.rank || 9) - (b.rank || 9)) ||
      (a.supplier < b.supplier ? -1 : a.supplier > b.supplier ? 1 : String(a.type || a.partNo) < String(b.type || b.partNo) ? -1 : 1);
  }
  function issueTitle(i) { return i.kind === 'part' ? '품번 ' + i.partNo + (i.partName ? '(' + i.partName + ')' : '') : i.type; }
  function issueText(i) { return i.kind === 'part' ? '동일품번 ' + i.partNo + ' ' + i.count + '건 [' + i.types.join(', ') + ']' : i.label + ' ' + i.type + ' ' + i.count + '건'; }
  /* 옛 이름(findIssues) — 화면·테스트가 기간 끝을 넘기던 모양 그대로 받습니다 */
  function findIssues(sel, to, rules) { return analyze(sel, {}, rules).issues; }

  /* ── 6. Master Plan 초안 ─────────────────────────────────────
     ① WORST 1·2·3 불량유형마다 사전의 기본 대책(유형마다 1주씩 늦춰 시작해 한꺼번에 몰리지 않게)
     ② 동일품번 2건 이상인 품번마다 「반복 불량 원인 분석·재발 방지」 한 줄
     진행상태(저장값): open = 진행중(마감이 지나면 날짜로 미완료) · done = 완료 · hold = 보류 */
  var STATUS = { open: '진행중', done: '완료', hold: '보류' };
  var idCounter = 0;
  function defaultId() { idCounter++; return 'p' + Date.now().toString(36) + idCounter.toString(36); }
  function makeRow(o, newId) {
    return { id: (newId || defaultId)(), supplier: o.supplier, type: o.type, action: o.action, dept: o.dept || '품질', owner: o.owner || '',
      planStart: o.planStart || '', planEnd: o.planEnd || '', actStart: '', actEnd: '', status: 'open', note: o.note || '', source: o.source || 'manual' };
  }
  function buildDraft(an, types, opts) {
    var supplier = opts.supplier, base = opts.baseDate, out = [];
    var byName = {}; (types || []).forEach(function (t) { byName[t.name] = t; });
    var mine = an.issues.filter(function (i) { return i.supplier === supplier; });
    var worst = mine.filter(function (i) { return i.kind === 'worst'; }).sort(function (a, b) { return a.rank - b.rank; });
    if (opts.maxTypes) worst = worst.slice(0, opts.maxTypes);
    worst.forEach(function (d, i) {
      var t = byName[d.type];
      var acts = t && t.actions && t.actions.length ? t.actions : [{ text: d.type + ' 원인 조사와 개선대책 수립', dept: '품질', start: 0, weeks: 2 }, { text: '유효성 확인(대책 후 재발 여부 추적)', dept: '품질', start: 2, weeks: 4 }];
      acts.forEach(function (a) {
        var ps = addDays(base, 7 * (a.start + i));
        out.push(makeRow({ supplier: supplier, type: d.type, action: a.text, dept: a.dept, planStart: ps, planEnd: addDays(ps, 7 * a.weeks - 1), source: 'template' }, opts.newId));
      });
    });
    mine.filter(function (i) { return i.kind === 'part'; }).forEach(function (p) {
      var ps = base;
      out.push(makeRow({ supplier: supplier, type: p.mainType,
        action: '동일품번 ' + p.partNo + (p.partName ? '(' + p.partName + ')' : '') + ' 반복 불량 ' + p.count + '건 — 원인 분석·재발 방지 대책', dept: '품질',
        planStart: ps, planEnd: addDays(ps, 13), source: 'template' }, opts.newId));
    });
    return out;
  }

  /* ── 7. 간트 칸·추진 집계 ────────────────────────────────────── */
  function rowDates(r) { return [r.planStart, r.planEnd, r.actStart, r.actEnd].filter(function (d) { return /^\d{4}-\d\d-\d\d$/.test(d || ''); }); }
  /* 진행상태(보이는 값) — 4차(2026-10-27) 수강생 요청: 진행중 · 완료 · 미완료 · 보류 네 가지만, 진행률은 쓰지 않음
     완료  : 완료로 둔 활동(실적 완료일 필요)
     보류  : 보류로 둔 활동
     미완료: 완료·보류가 아니고 개선 마감(계획 완료일)이 오늘보다 앞 — 마감이 지났는데 끝나지 않음
     진행중: 완료·보류가 아니고 마감이 아직 안 지남(개선 시작~마감 사이. 시작 전 활동과 날짜가 빈 활동도 여기에 셉니다)
     3차의 지연 = 미완료, 진행 중 · 예정 = 진행중 으로 바뀌었습니다(저장값은 그대로라 옮길 것이 없음). */
  var PHASE = { doing: '진행중', done: '완료', late: '미완료', hold: '보류' };
  var PHASE_ORDER = ['doing', 'done', 'late', 'hold'];
  function phaseOf(r, today) {
    if (r.status === 'done') return 'done';
    if (r.status === 'hold') return 'hold';
    if (r.planEnd && r.planEnd < today) return 'late';
    return 'doing';
  }
  /* 진행중 가운데 개선 시작 전인 활동 — 화면에 「시작 전」으로 작게 덧붙일 때만 씁니다 */
  function notStarted(r, today) { return phaseOf(r, today) === 'doing' && !r.actStart && !(r.planStart && r.planStart <= today); }
  function isOverdue(r, today) { return phaseOf(r, today) === 'late'; }
  function planStats(rows, today) {
    var s = { total: rows.length, done: 0, open: 0, hold: 0, overdue: 0, doing: 0, late: 0, notStarted: 0, start: '', end: '', days: 0, rate: 0 };
    rows.forEach(function (r) {
      s[r.status in STATUS ? r.status : 'open']++;
      var ph = phaseOf(r, today);
      if (ph === 'doing' || ph === 'late') s[ph]++;
      if (notStarted(r, today)) s.notStarted++;
      if (r.planStart && (!s.start || r.planStart < s.start)) s.start = r.planStart;
      if (r.planEnd && (!s.end || r.planEnd > s.end)) s.end = r.planEnd;
    });
    s.overdue = s.late;
    if (s.start && s.end && s.end >= s.start) s.days = daysBetween(s.start, s.end) + 1;
    s.rate = s.total ? Math.round(s.done / s.total * 100) : 0;
    return s;
  }
  /* scale = 'week'(주차 칸, 위 줄 = 월) | 'month'(월 칸, 위 줄 = 연도) */
  function ganttColumns(rows, scale, opts) {
    opts = opts || {};
    var ds = [];
    rows.forEach(function (r) { ds = ds.concat(rowDates(r)); });
    if (opts.today) ds.push(opts.today);
    ds.sort();
    var first = ds[0] || opts.baseDate || opts.today, last = ds[ds.length - 1] || first;
    if (!first) return [];
    var cols = [], min = opts.minCols || (scale === 'month' ? 6 : 10);
    if (scale === 'month') {
      for (var m = monthOf(first); cols.length < 120 && (m <= monthOf(last) || cols.length < min); m = monthAdd(m, 1)) {
        cols.push({ start: m + '-01', end: monthLast(m), label: +m.slice(5) + '월', group: m.slice(0, 4) + '년', full: m.slice(0, 4) + '년 ' + (+m.slice(5)) + '월' });
      }
    } else {
      for (var w = mondayOf(first); cols.length < 260 && (w <= last || cols.length < min); w = addDays(w, 7)) {
        var wi = weekInfo(w);
        cols.push({ start: w, end: wi.sunday, label: wi.week + '주', group: wi.year + '년 ' + wi.month + '월', full: wi.label });
      }
    }
    return cols;
  }
  /* 기간이 걸치는 칸 번호 [처음, 끝] — 없으면 null. 끝이 비면 시작 칸 하나 */
  function span(start, end, cols) {
    if (!start || !cols.length) return null;
    end = end && end >= start ? end : start;
    var a = -1, b = -1;
    for (var i = 0; i < cols.length; i++) {
      if (a < 0 && cols[i].end >= start) a = i;
      if (cols[i].start <= end) b = i;
    }
    if (a < 0 || b < 0 || b < a) return null;
    return [a, b];
  }
  function groupSpans(cols) {
    var out = [];
    cols.forEach(function (c) { var g = out[out.length - 1]; if (g && g.group === c.group) g.n++; else out.push({ group: c.group, n: 1 }); });
    return out;
  }
  function cleanPlanRow(r) {
    var o = makeRow({ supplier: str(r.supplier), type: str(r.type), action: str(r.action), dept: str(r.dept), owner: str(r.owner), note: str(r.note), source: r.source === 'template' || r.source === 'ai' || r.source === 'reply' ? r.source : 'manual' }, function () { return str(r.id) || defaultId(); });
    ['planStart', 'planEnd', 'actStart', 'actEnd'].forEach(function (k) { o[k] = parseDate(r[k]); });
    o.status = r.status in STATUS ? r.status : 'open';
    if (str(r.replyAt)) o.replyAt = str(r.replyAt);
    return o;
  }
  /* 입력 점검 — 화면의 오른쪽 칸에서 저장 전에 */
  function checkPlanRow(r) {
    var e = [];
    if (!str(r.action)) e.push('개선대책을 적어 주세요.');
    if (r.planStart && r.planEnd && r.planEnd < r.planStart) e.push('계획 완료일이 시작일보다 빠릅니다.');
    if (r.actStart && r.actEnd && r.actEnd < r.actStart) e.push('실적 완료일이 시작일보다 빠릅니다.');
    if (r.status === 'done' && !r.actEnd) e.push('완료로 두려면 실적 완료일을 적어 주세요.');
    return e;
  }

  /* ── 8. 계획 확정 ──────────────────────────────────────────── */
  function planKey(rows) { return JSON.stringify(rows.map(function (r) { return [r.type, r.action, r.dept, r.owner, r.planStart, r.planEnd]; })); }
  function confirmPlan(plan, at) {
    plan.confirmed = plan.confirmed || [];
    var snap = { version: plan.confirmed.length + 1, at: at, rows: clone(plan.rows) };
    plan.confirmed.push(snap);
    return snap;
  }
  function lastConfirmed(plan) { return plan && plan.confirmed && plan.confirmed.length ? plan.confirmed[plan.confirmed.length - 1] : null; }
  /* 확정 뒤 계획(대책·부서·담당자·계획일)이 바뀌었는지 — 실적·상태 입력은 바뀜으로 보지 않습니다 */
  function planChangedSince(plan) { var c = lastConfirmed(plan); return !!c && planKey(c.rows) !== planKey(plan.rows); }
  function planStateText(plan) {
    var c = lastConfirmed(plan);
    if (!c) return '미확정 초안';
    return '확정 v' + c.version + ' (' + c.at.replace('T', ' ').slice(0, 16) + ')' + (planChangedSince(plan) ? ' 이후 계획 수정 있음' : '');
  }

  /* ── 9. AI 프롬프트·답 읽기 ─────────────────────────────────
     협력사 이름은 「업체1」, 품번은 「품목1」 같은 별칭으로 바꿉니다(기본). 대응표는 이 브라우저에만 둡니다. */
  var ALIAS = { suppliers: '업체', parts: '품목' };
  /* 붙여 넣을 AI — 사내 허용 Copilot, 임시 채팅(대화 기록을 남기지 않는 창)으로 ChatGPT · Claude · Gemini */
  var AI_TARGETS = [
    { id: 'copilot', name: 'Copilot', how: '사내에서 허용된 Copilot 창을 열어 붙여 넣습니다.' },
    { id: 'chatgpt', name: 'ChatGPT (임시 채팅)', how: 'ChatGPT 에서 「임시 채팅」을 켠 뒤 붙여 넣습니다.' },
    { id: 'claude', name: 'Claude (임시 채팅)', how: 'Claude 에서 「임시 채팅(Incognito)」을 켠 뒤 붙여 넣습니다.' },
    { id: 'gemini', name: 'Gemini (임시 채팅)', how: 'Gemini 에서 「임시 채팅」을 켠 뒤 붙여 넣습니다.' }
  ];
  function assignAliases(map, names, prefix) {
    var out = {}, used = 0, re = new RegExp('^' + prefix + '(\\d+)$');
    Object.keys(map || {}).forEach(function (k) { var m = re.exec(map[k]); if (m) { out[k] = map[k]; used = Math.max(used, +m[1]); } });
    uniq((names || []).map(str).filter(Boolean)).forEach(function (k) { if (!out[k]) out[k] = prefix + (++used); });
    return out;
  }
  function updateAliases(aliases, records) {
    aliases = aliases || {};
    return {
      suppliers: assignAliases(aliases.suppliers, records.map(function (r) { return r.supplier; }).sort(), ALIAS.suppliers),
      parts: assignAliases(aliases.parts, records.map(function (r) { return r.partNo; }).sort(), ALIAS.parts)
    };
  }
  function escRe(s) { return s.replace(/[.*+?^${}()|[\]\\\/-]/g, '\\$&'); }
  /* 글 속에서 바꿀 수 있는 이름 — 두 글자 이상(한 글자는 다른 낱말과 헷갈림) */
  function nameRegex(names) {
    var list = names.filter(function (n) { return n.length >= 2; }).sort(function (a, b) { return b.length - a.length; }).map(escRe);
    return list.length ? new RegExp('(^|[^0-9A-Za-z])(' + list.join('|') + ')(?![0-9A-Za-z])', 'g') : null;
  }
  function maskText(text, aliases) {
    var s = text == null ? '' : String(text);
    var all = {};
    ['suppliers', 'parts'].forEach(function (k) { var m = (aliases && aliases[k]) || {}; Object.keys(m).forEach(function (n) { all[n] = m[n]; }); });
    var re = nameRegex(Object.keys(all));
    return re ? s.replace(re, function (x, pre, n) { return pre + all[n]; }) : s;
  }
  /* AI 답 → 원래 이름. 「업체3」「품목3」만 바꾸고 「품목 3개」「품목3개」처럼 수를 세는 말은 그대로 둡니다 */
  function unmaskText(text, aliases) {
    var back = {};
    ['suppliers', 'parts'].forEach(function (k) { var m = (aliases && aliases[k]) || {}; Object.keys(m).forEach(function (n) { back[m[n]] = n; }); });
    return String(text == null ? '' : text).replace(new RegExp('(' + ALIAS.suppliers + '|' + ALIAS.parts + ')(\\d+)(?![0-9개건종가])', 'g'), function (all) { return back[all] || all; });
  }
  function findLeaks(text, aliases) {
    var s = String(text || ''), names = [];
    ['suppliers', 'parts'].forEach(function (k) { names = names.concat(Object.keys((aliases && aliases[k]) || {})); });
    return names.filter(function (n) { var re = nameRegex([n]); return re && re.test(s); });
  }
  /* ctx = { supplier, an(analyze 결과, 협력사로 거른 것), rows(지금 계획), types, depts, baseDate, aliases(null 이면 가리지 않음), compact(짧게) }
     사내 허용 AI = Copilot, 임시 채팅으로 ChatGPT·Claude·Gemini (2026-10-27 수강생 답) — API 키 없이 복사·붙여 넣기로 씁니다.
     compact: 입력 글자 수가 빠듯한 채팅 창을 위해 문제점 기록을 유형마다 2건으로, 지금 초안 줄을 뺍니다. */
  function buildAiPromptInfo(ctx) {
    var al = ctx.aliases || null, an = ctx.an, L = [], cp = !!ctx.compact;
    function M(s) { return al ? maskText(s, al) : String(s == null ? '' : s); }
    L.push('너는 모기업 품질담당자가 협력사 품질개선 계획(MASTER PLAN)을 세우는 것을 돕는 보조야. 아래 [사실]은 협력사 공정불량 기록에서 도구가 계산한 값이야.');
    L.push('[사실]에 있는 숫자와 내용만 근거로 삼아줘. 숫자를 새로 만들거나 바꾸지 말고, 기록에 없는 원인은 반드시 「추정」이라고 표시해줘.');
    if (al) L.push('협력사 이름과 품번은 보안상 「' + ALIAS.suppliers + '1」「' + ALIAS.parts + '1」 같은 별칭으로 바꿔 두었어. 답에서도 별칭을 띄어 쓰지 말고 그대로 써줘(실제 이름을 짐작하지 말아줘).');
    L.push('');
    L.push('[사실]');
    L.push('- 협력사: ' + M(ctx.supplier));
    L.push('- 분석 기간: ' + an.from + ' ~ ' + an.to + ', 공정불량 ' + an.total + '건' + (an.unclassified ? ' (유형 미분류 ' + an.unclassified + '건 포함)' : ''));
    L.push('- 불량유형별 건수: ' + (an.byType.map(function (t) { return t.type + ' ' + t.count + '건(' + t.pct + '%)'; }).join(', ') || '없음'));
    var iss = an.issues.filter(function (i) { return i.supplier === ctx.supplier; });
    var w = iss.filter(function (i) { return i.kind === 'worst'; }).sort(function (a, b) { return a.rank - b.rank; });
    var pp = iss.filter(function (i) { return i.kind === 'part'; });
    L.push('- 우선 개선 항목 ① 불량유형 WORST: ' + (w.map(function (i) { return i.label + ' ' + i.type + ' ' + i.count + '건'; }).join(', ') || '없음'));
    L.push('- 우선 개선 항목 ② 동일품번 2건 이상: ' + (pp.slice(0, 10).map(function (p) { return M(p.partNo) + (p.partName ? '(' + p.partName + ')' : '') + ' ' + p.count + '건 [' + p.types.join(', ') + ']'; }).join('; ') || '없음'));
    var byType = {}, per = cp ? 2 : 4;
    an.rows.forEach(function (r) { if (r.supplier !== ctx.supplier) return; var a = byType[r.type] || (byType[r.type] = []); if (a.length < per && a.indexOf(r.problem) < 0) a.push(r.problem); });
    var tl = Object.keys(byType);
    if (tl.length) {
      L.push('- 유형별 문제점 기록(원문 일부):');
      tl.forEach(function (t) { L.push('  · ' + t + ': ' + byType[t].map(function (p) { return '「' + M(p) + '」'; }).join(' ')); });
    }
    L.push('- 담당부서(협력사 내부 부서) 목록: ' + (ctx.depts || DEFAULT_DEPTS).join(', '));
    L.push('- 계획 착수 기준일: ' + ctx.baseDate);
    if (!cp && ctx.rows && ctx.rows.length) {
      L.push('- 지금 초안(기본 대책 사전에서 만든 것):');
      ctx.rows.forEach(function (r) { L.push('  · ' + r.type + ' | ' + M(r.action) + ' | ' + r.dept + ' | ' + r.planStart + '~' + r.planEnd); });
    }
    L.push('');
    L.push('[요청]');
    L.push('1. 첫 줄에 「요약:」으로 시작하는 한 줄 요약을 써줘.');
    L.push('2. 우선 개선 항목의 불량유형마다 구체적인 개선대책 2~4개를 아래 표 형식으로만 써줘. 담당부서는 위 목록에서 하나만 골라줘.');
    L.push('| 불량유형 | 개선대책 | 담당부서 | 시작(기준일부터 몇 주 뒤) | 기간(주) |');
    L.push('|---|---|---|---|---|');
    L.push('3. 표 아래에 「확인 필요:」로 시작해 현장에서 확인할 것을 한 줄씩 적어줘.');
    var text = L.join('\n');
    return { prompt: text, leaks: al ? findLeaks(text, al) : [] };
  }
  /* AI 답의 담당부서 → 목록의 이름(「품질팀」「생산기술팀」 → 품질·생산기술). 맞는 것이 없으면 그대로 */
  function matchDept(s, depts) {
    var v = str(s), n = v.replace(/\s+/g, '').replace(/(팀|부|부서|파트|그룹)$/, '');
    if (!v) return '';
    var list = depts || DEFAULT_DEPTS;
    for (var i = 0; i < list.length; i++) if (list[i] === v || list[i].replace(/\s+/g, '') === n) return list[i];
    for (i = 0; i < list.length; i++) if (n.length >= 2 && list[i].length >= 2 && (n.indexOf(list[i]) >= 0 || list[i].indexOf(n) >= 0)) return list[i];
    return v;
  }
  function matchType(s, types) {
    if (!str(s)) return '';
    return findTypeName(s, types) || str(s);
  }
  function weeksNum(s, def) { var m = /(\d+)/.exec(str(s)); return m ? +m[1] : def; }
  /* AI 답 읽기 — 「| 유형 | 대책 | 부서 | 시작 | 기간 |」 표(또는 탭으로 나뉜 줄). opts = { types, aliases, supplier, baseDate, newId } */
  function parseAiPlan(text, opts) {
    var t = unmaskText(text, opts.aliases), rows = [], skipped = 0, summary = '', checks = [], inChecks = false;
    t.split(/\r?\n/).forEach(function (line) {
      var s = line.trim();
      if (!s) return;
      var sm = /^[#\s]*\**\s*요약\s*\**\s*[:：]\s*\**\s*(.+)$/.exec(s);
      if (sm && !summary) { summary = sm[1].replace(/\*+$/, '').trim(); return; }
      var cm = /^[#\s]*\**\s*확인\s*필요\s*\**\s*[:：]\s*\**\s*(.*)$/.exec(s);
      if (cm) { inChecks = true; if (cm[1].trim()) checks.push(cm[1].trim()); return; }
      var cells = null;
      if (s.charAt(0) === '|') cells = s.replace(/^\|/, '').replace(/\|$/, '').split('|').map(str);
      else if (s.indexOf('\t') >= 0) cells = s.split('\t').map(str);
      if (!cells) { if (inChecks) checks.push(s.replace(/^[-*·\d.)\s]+/, '')); return; }
      inChecks = false;
      if (cells.every(function (c) { return /^:?-{2,}:?$/.test(c) || !c; })) return;
      if (/불량\s*유형/.test(cells[0]) && /대책/.test(cells[1] || '')) return;
      if (cells.length < 2 || !cells[1]) { skipped++; return; }
      var st = weeksNum(cells[3], 0), wk = Math.max(1, weeksNum(cells[4], 2));
      var ps = addDays(opts.baseDate, 7 * st);
      rows.push(makeRow({ supplier: opts.supplier, type: matchType(cells[0], opts.types) || '기타', action: cells[1].replace(/<br\s*\/?>/gi, ' '), dept: matchDept(cells[2], opts.depts) || '품질',
        planStart: ps, planEnd: addDays(ps, 7 * wk - 1), source: 'ai' }, opts.newId));
    });
    return { summary: summary, rows: rows, skipped: skipped, checks: checks };
  }

  /* ── 10. Excel 추출 = 「불량목록」 · 「MASTER PLAN」 두 시트 (2026-10-27 수강생 요청) ─────────────
     MASTER PLAN 시트는 협력사가 고쳐 회신하는 양식입니다. 맨 왼쪽 ID 칸과 맨 오른쪽 기준 칸은 숨겨 두고,
     회신 파일을 「진행사항 업로드」로 다시 읽을 때 ID 로 행을 맞추고 기준 칸(내보낼 때의 값)으로 누가 바꿨는지 가립니다.
     meta 는 js/xlsx-out.js 가 머리글 굵게·칸 색·틀 고정·거르기·숨김 열·고르기 목록을 붙일 때 씁니다.
     ctx = { suppliers:[이름], records(enrich 된 전체), plans, rules, today, types, depts } */
  var REPORT_TITLE = '협력사 공정불량 개선 MASTER PLAN';
  var SHEET_LIST = '불량목록', SHEET_PLAN = 'MASTER PLAN';
  function sourceText(r) { return r.source === 'ai' ? 'AI 답 반영' : r.source === 'template' ? '기본 대책' : r.source === 'reply' ? '협력사 회신' : '직접 입력'; }
  /* 회신 양식의 칸 — k: 계획 행의 필드, edit: 협력사가 고치는 칸 */
  var PCOLS = [
    { k: 'id', head: 'ID', heads: ['id'], w: 6 },
    { k: 'no', head: 'No', heads: ['no', '번호', '순번'], w: 5 },
    { k: 'supplier', head: '협력사', heads: ['협력사', '협력사명', '업체', '업체명'], w: 12 },
    { k: 'type', head: '불량유형', heads: ['불량유형', '유형'], w: 12, edit: true },
    { k: 'action', head: '개선대책', heads: ['개선대책', '대책', '개선과제', '개선활동'], w: 44, edit: true },
    { k: 'dept', head: '담당부서', heads: ['담당부서', '부서'], w: 10, edit: true },
    { k: 'owner', head: '담당자', heads: ['담당자'], w: 10, edit: true },
    { k: 'planStart', head: '개선 시작(계획)', heads: ['개선시작', '계획시작', '시작일'], w: 12, edit: true },
    { k: 'planEnd', head: '개선 마감(계획)', heads: ['개선마감', '계획완료', '계획마감', '계획종료', '마감일'], w: 12, edit: true },
    { k: 'actStart', head: '실적 시작', heads: ['실적시작'], w: 11, edit: true },
    { k: 'actEnd', head: '실적 완료', heads: ['실적완료', '완료일'], w: 11, edit: true },
    { k: 'status', head: '진행상태', heads: ['진행상태', '상태', '실적상태'], w: 9, edit: true },
    { k: 'note', head: '비고', heads: ['비고', '메모'], w: 30, edit: true },
    { k: 'base', head: '기준(지우지 마세요)', heads: ['기준'], w: 8 }
  ];
  var PFIELDS = PCOLS.filter(function (c) { return c.edit; });
  var PHASE_IN = ['진행중', '완료', '미완료', '보류'];
  var COL_STATUS = PCOLS.map(function (c) { return c.k; }).indexOf('status');
  /* 비교에 쓰는 값 — 내보낼 때·기준 칸·지금 값 모두 이 모양으로 맞춰 비교합니다 */
  function curVal(r, k) {
    if (k === 'status') return r.status in STATUS ? r.status : 'open';
    return r[k] == null ? '' : str(r[k]);
  }
  function baseOf(r) { var o = {}; PFIELDS.forEach(function (c) { o[c.k] = curVal(r, c.k); }); return o; }
  function planSheetRows(rows, t) {
    return rows.map(function (r, n) {
      var b = baseOf(r);
      return [r.id, n + 1, r.supplier, r.type, r.action, r.dept, r.owner, r.planStart, r.planEnd, r.actStart, r.actEnd, PHASE[phaseOf(r, t)], r.note, JSON.stringify(b)];
    });
  }
  function reportBook(ctx) {
    var sheets = {}, meta = {};
    var t = ctx.today, rules = cleanRules(ctx.rules);
    var per = ctx.suppliers.map(function (s) {
      var plan = ctx.plans[s] || { rows: [], confirmed: [] };
      return { s: s, plan: plan, an: analyze(ctx.records, { supplier: s, from: plan.from, to: plan.to }, rules) };
    });
    var rows = []; per.forEach(function (x) { rows = rows.concat(x.an.rows.filter(function (r) { return campusMatch(r, ctx.campus); })); });
    rows.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : (a.supplier < b.supplier ? -1 : a.supplier > b.supplier ? 1 : 0); });

    // 1. 불량목록 — 불량 등록 화면의 칸 순서
    // 캠퍼스 칸(4차) — 울산 · 인천 RAW DATA 를 합쳐도 어디서 온 행인지 보이게. 관리번호 = 각 시스템의 문제점 번호 · QM noti No.
    var list = [['No', '캠퍼스', '발생일자', '협력사', '품번', '품명', '불량유형', '내용', '수량', '조치', '관리번호']];
    rows.forEach(function (r, i) { list.push([i + 1, r.campus || '', r.date, r.supplier, r.partNo, r.partName, r.type, r.problem, typeof r.qty === 'number' ? r.qty : '', r.action || '', r.srcNo || '']); });
    sheets[SHEET_LIST] = list;
    meta[SHEET_LIST] = { head: [0], freeze: { rows: 1 }, filter: 'A1:K' + list.length, widths: [6, 8, 11, 12, 12, 14, 12, 40, 7, 30, 14] };

    // 2. MASTER PLAN — 협력사가 고쳐 회신하는 양식
    var all = []; per.forEach(function (x) { (x.plan.rows || []).forEach(function (r) { all.push(r); }); });
    var st = planStats(all, t);
    var mp = [[REPORT_TITLE + ' — ' + (ctx.suppliers.length === 1 ? ctx.suppliers[0] : '전체 ' + ctx.suppliers.length + '곳')],
      ['기준일 ' + t + ' · 전체 ' + st.total + ' · 진행중 ' + st.doing + ' · 완료 ' + st.done + ' · 미완료 ' + st.late + ' · 보류 ' + st.hold +
        (ctx.suppliers.length === 1 && all.length ? ' · 계획 상태: ' + planStateText(per[0].plan) : '')],
      ['노란 머리 칸을 고쳐 회신해 주세요(진행상태: 진행중 · 완료 · 미완료 · 보류 — 완료는 실적 완료일도 함께). 숨긴 ID 칸과 기준 칸은 지우거나 바꾸지 마세요. 새 활동은 표 아래 빈 줄에 ID 없이 적으면 됩니다.'], []];
    var H = mp.length;
    mp.push(PCOLS.map(function (c) { return c.head; }));
    var cells = [];
    PCOLS.forEach(function (c, i) { if (c.edit) cells.push({ r: H, c: i, style: 'edit' }); });
    planSheetRows(all, t).forEach(function (line) { mp.push(line); });
    var last = Math.max(mp.length, H + 1), lastCol = colLetter(PCOLS.length - 1), stCol = colLetter(COL_STATUS);
    var stRange = stCol + (H + 2) + ':' + stCol + (H + 500);
    sheets[SHEET_PLAN] = mp;
    /* 진행상태 음영(4차) = 엑셀 「조건부 서식」 — 협력사가 값을 고르면 색이 바로 바뀝니다. 완료 초록 · 진행중 주황 · 미완료 빨강 · 보류 회색 */
    meta[SHEET_PLAN] = { title: [0], head: [H], cells: cells, freeze: { rows: H + 1 }, filter: 'A' + (H + 1) + ':' + lastCol + last,
      widths: PCOLS.map(function (c) { return c.w; }), hidden: [0, PCOLS.length - 1],
      cf: [{ sqref: stRange, rules: [{ equal: PHASE.done, style: 'done' }, { equal: PHASE.doing, style: 'doing' }, { equal: PHASE.late, style: 'late' }, { equal: PHASE.hold, style: 'hold' }] }],
      lists: [{ sqref: stRange, items: PHASE_IN }, { sqref: 'F' + (H + 2) + ':F' + (H + 500), items: (ctx.depts || DEFAULT_DEPTS).slice(0, 20), soft: true }] };
    return { sheets: sheets, meta: meta };
  }
  function reportSheets(ctx) { return reportBook(ctx).sheets; }
  function colLetter(c) { var s = ''; c++; while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; }
  /* 엑셀 백업 = 불량 목록을 붙여넣기 양식 순서 그대로(다시 「엑셀 복원」·붙여넣기로 읽힘) */
  function backupSheet(records) {
    var aoa = [PASTE_ORDER.map(fieldLabel).concat([fieldLabel('srcNo'), fieldLabel('campus')])];
    records.forEach(function (r) { aoa.push([r.date, r.supplier, r.partNo, r.partName, r.typeManual || r.typeIn || '', r.problem, typeof r.qty === 'number' ? r.qty : '', r.action || '', r.srcNo || '', r.campus || '']); });
    return aoa;
  }

  /* ── 10-2. 진행사항 업로드 — 협력사가 고친 MASTER PLAN 엑셀을 다시 읽어 바뀐 칸만 반영 ─────────────
     행 맞추기: ① 숨긴 ID 칸 ② ID 가 없거나 모르는 ID 면 같은 협력사·같은 개선대책 글
     칸마다: 파일 값 = 지금 값 → 그대로 / 파일 값 = 기준 값(내보낼 때) → 협력사는 안 고침(이 도구에서 고친 값 유지)
             지금 값 = 기준 값 → 협력사가 고침(반영) / 셋이 모두 다름 → 충돌(기본은 지금 값 유지, 고르면 파일 값)
     같은 파일을 두 번 올려도 두 번째는 「바뀐 것 없음」이 됩니다(새 줄도 개선대책 글로 맞춰 다시 더하지 않음). */
  function parseProgressSheet(aoa, types) {
    var hr = -1, map = {};
    for (var r = 0; r < Math.min(30, aoa.length) && hr < 0; r++) {
      var row = aoa[r] || [], m = {};
      row.forEach(function (v, c) {
        var n = normHead(v); if (!n) return;
        PCOLS.forEach(function (col) { if (m[col.k] == null && col.heads.some(function (h) { return n === h || (h.length >= 2 && n.indexOf(h) === 0); })) m[col.k] = c; });
      });
      if (m.action != null && (m.id != null || m.status != null || m.planEnd != null)) { hr = r; map = m; }
    }
    if (hr < 0) return { ok: false, rows: [], error: '「개선대책」 칸이 있는 머리 줄을 찾지 못했습니다. 이 도구에서 내려받은 MASTER PLAN 엑셀인지 확인해 주세요.' };
    var out = [];
    for (r = hr + 1; r < aoa.length; r++) {
      var line = aoa[r] || [];
      var get = function (k) { return map[k] == null ? undefined : line[map[k]]; };
      var x = { line: r + 1, id: str(get('id')), supplier: str(get('supplier')), vals: {}, bad: [] };
      PFIELDS.forEach(function (c) { if (map[c.k] != null) x.vals[c.k] = get(c.k); });
      var any = PFIELDS.some(function (c) { return str(x.vals[c.k]) !== ''; });
      if (!x.id && !str(x.vals.action)) { if (any && line.some(function (v) { return str(v); })) { x.bad.push('개선대책이 비어 있어 읽지 않았습니다'); out.push(x); } continue; }
      var b = str(get('base'));
      if (b) { try { var o = JSON.parse(b); if (o && typeof o === 'object') x.base = o; } catch (e) { x.base = null; } }
      out.push(x);
    }
    return { ok: true, headerRow: hr, map: map, rows: out, hasId: map.id != null, hasBase: map.base != null };
  }
  function normField(k, v, types) {
    var s = v == null ? '' : (v instanceof Date ? v : typeof v === 'number' ? v : str(v));
    if (k === 'planStart' || k === 'planEnd' || k === 'actStart' || k === 'actEnd') { if (s === '' ) return { v: '' }; var d = parseDate(s); return d ? { v: d } : { v: '', bad: '날짜를 읽지 못했습니다(' + str(v) + ')' }; }
    /* 진행상태 글 → 저장값. 진행중 · 미완료(날짜로 정해짐)와 3차 값 진행 중 · 지연 · 예정은 모두 open */
    if (k === 'status') { var t = str(s).replace(/\s+/g, ''); return { v: /미완료/.test(t) ? 'open' : /완료/.test(t) ? 'done' : /보류/.test(t) ? 'hold' : 'open' }; }
    if (k === 'type') return { v: findTypeName(s, types) || str(s) };
    return { v: str(s).replace(/\r\n?/g, '\n') };
  }
  function showVal(k, v) { return k === 'status' ? (v === 'open' ? '진행중·미완료' : STATUS[v] || v) : v; }
  function fieldLabelOf(k) { var c = PCOLS.filter(function (x) { return x.k === k; })[0]; return c ? c.head.replace(/\(.*\)$/, '') : k; }
  function actionKey(s) { return str(s).replace(/\s+/g, '').toLowerCase(); }
  /* opts = { types, today, knownSuppliers:[이름] } → { items:[…], counts, notInFile:[…], suppliers } */
  function diffProgress(plans, parsed, opts) {
    opts = opts || {};
    var byId = {}, items = [], seen = {}, sups = {};
    Object.keys(plans || {}).forEach(function (s) { (plans[s].rows || []).forEach(function (r) { byId[r.id] = { s: s, r: r }; }); });
    parsed.rows.forEach(function (x) { if (x.supplier) sups[x.supplier] = 1; });
    var fileSups = Object.keys(sups), only = fileSups.length === 1 ? fileSups[0] : '';
    var known = {}; (opts.knownSuppliers || []).concat(Object.keys(plans || {})).forEach(function (s) { known[s] = 1; });
    var touched = {};
    parsed.rows.forEach(function (x) {
      var it = { line: x.line, id: x.id, supplier: x.supplier || only, action: str(x.vals.action), kind: '', changes: [], conflicts: [], reasons: x.bad.slice() };
      if (x.bad.length) { it.kind = 'invalid'; items.push(it); return; }
      if (x.id && seen[x.id]) { it.kind = 'invalid'; it.reasons.push('같은 ID 가 ' + seen[x.id] + '번째 줄에도 있습니다'); items.push(it); return; }
      if (x.id) seen[x.id] = x.line;
      var hit = x.id && byId[x.id] ? byId[x.id] : null;
      if (hit && x.supplier && x.supplier !== hit.s) hit = null;
      if (!hit) {
        var s0 = it.supplier, ak = actionKey(x.vals.action);
        if (s0 && plans[s0] && ak) { var m = plans[s0].rows.filter(function (r) { return actionKey(r.action) === ak; })[0]; if (m) { hit = { s: s0, r: m }; it.byText = true; } }
      }
      var vals = {}, bad = [];
      PFIELDS.forEach(function (c) { if (c.k in x.vals) { var n = normField(c.k, x.vals[c.k], opts.types); vals[c.k] = n.v; if (n.bad) bad.push(fieldLabelOf(c.k) + ': ' + n.bad); } });
      if (bad.length) { it.kind = 'invalid'; it.reasons = it.reasons.concat(bad); if (hit) { it.supplier = hit.s; it.target = hit.r.id; } items.push(it); return; }
      if (!hit) {
        if (x.id) { it.kind = 'unknown'; it.reasons.push('이 도구에 없는 ID 입니다(지운 활동이거나 다른 PC 에서 만든 계획)'); }
        else if (!it.supplier || !known[it.supplier]) { it.kind = 'unknown'; it.reasons.push(it.supplier ? '이 도구에 없는 협력사입니다' : '협력사 칸이 비어 있습니다'); }
        else it.kind = 'new';
        it.vals = vals;
        if (it.kind === 'new') { var nr = makeRow({ supplier: it.supplier, type: vals.type || '기타', action: vals.action }, function () { return 'tmp'; }); PFIELDS.forEach(function (c) { if (c.k in vals) nr[c.k] = vals[c.k]; });
          var e0 = checkPlanRow(nr); if (e0.length) { it.kind = 'invalid'; it.reasons = it.reasons.concat(e0); } }
        items.push(it); return;
      }
      it.supplier = hit.s; it.target = hit.r.id; touched[hit.r.id] = 1;
      var base = x.base && typeof x.base === 'object' ? x.base : null;
      var next = clone(hit.r);
      PFIELDS.forEach(function (c) {
        if (!(c.k in vals)) return;
        var f = vals[c.k], cur = curVal(hit.r, c.k), b = base && c.k in base ? base[c.k] : cur;
        if (String(f) === String(cur)) return;
        if (String(f) === String(b)) return;                    // 협력사는 안 고침 — 이 도구에서 고친 값 유지
        var ch = { k: c.k, label: fieldLabelOf(c.k), from: cur, to: f, show: [showVal(c.k, cur), showVal(c.k, f)] };
        if (String(cur) === String(b)) { it.changes.push(ch); next[c.k] = f; }
        else { ch.base = b; it.conflicts.push(ch); }
      });
      var errs = checkPlanRow(next);
      if (errs.length && (it.changes.length || it.conflicts.length)) { it.kind = 'invalid'; it.reasons = it.reasons.concat(errs); }
      else it.kind = it.changes.length || it.conflicts.length ? 'changed' : 'same';
      items.push(it);
    });
    var counts = { changed: 0, same: 0, new: 0, unknown: 0, invalid: 0, conflicts: 0, fields: 0 };
    items.forEach(function (it) { counts[it.kind]++; counts.conflicts += it.conflicts.length; counts.fields += it.changes.length; });
    var notInFile = [];
    fileSups.forEach(function (s) { if (plans[s]) plans[s].rows.forEach(function (r) { if (!touched[r.id]) notInFile.push({ supplier: s, id: r.id, action: r.action }); }); });
    return { items: items, counts: counts, notInFile: notInFile, suppliers: fileSups };
  }
  /* 반영 — take = { 줄번호: { 필드: true } } (충돌 칸 가운데 파일 값을 고른 것). 바뀐 행에는 replyAt 을 남깁니다 */
  function applyProgress(plans, diff, opts) {
    opts = opts || {};
    var take = opts.take || {}, at = opts.at || '', res = { changed: 0, added: 0, fields: 0, skipped: 0, details: [] };
    var byId = {}; Object.keys(plans).forEach(function (s) { plans[s].rows.forEach(function (r) { byId[r.id] = r; }); });
    diff.items.forEach(function (it) {
      if (it.kind === 'changed') {
        var r = byId[it.target]; if (!r) { res.skipped++; return; }
        var next = clone(r), list = it.changes.slice();
        it.conflicts.forEach(function (c) { if (take[it.line] && take[it.line][c.k]) list.push(c); });
        if (!list.length) return;
        list.forEach(function (c) { next[c.k] = c.to; });
        if (checkPlanRow(next).length) { res.skipped++; return; }
        list.forEach(function (c) { r[c.k] = c.to; });
        r.replyAt = at; res.changed++; res.fields += list.length;
        res.details.push({ supplier: it.supplier, id: r.id, action: r.action, changes: list.map(function (c) { return [c.label, c.show[0], c.show[1]]; }) });
      } else if (it.kind === 'new') {
        var plan = plans[it.supplier] || (plans[it.supplier] = { baseDate: opts.today || '', from: '', to: '', rows: [], confirmed: [], aiSummary: '', aiChecks: [], aiAnswer: '' });
        var nr = makeRow({ supplier: it.supplier, type: it.vals.type || '기타', action: it.vals.action, source: 'reply' }, opts.newId);
        PFIELDS.forEach(function (c) { if (c.k in it.vals && c.k !== 'type') nr[c.k] = it.vals[c.k]; });
        if (!nr.dept) nr.dept = '품질';
        nr.replyAt = at; plan.rows.push(nr); res.added++;
        res.details.push({ supplier: it.supplier, id: nr.id, action: nr.action, changes: [['새 활동', '', nr.action]] });
      }
    });
    return res;
  }
  /* 파일 지문(FNV-1a) — 같은 파일을 다시 올렸는지 알려 줄 때만 씁니다 */
  function fingerprint(u8) { var h = 0x811c9dc5; for (var i = 0; i < u8.length; i++) { h ^= u8[i]; h = Math.imul(h, 0x01000193) >>> 0; } return ('0000000' + h.toString(16)).slice(-8) + '-' + u8.length; }

  /* ── 11. 저장 형식 ─────────────────────────────────────────── */
  var SCHEMA_VERSION = 4; // 4 = 캠퍼스(campus · srcNo · src) · 진행률 없음
  /* 1단계(v1) 기본 사전 9종 → 수강생 8종. 사전 이름을 손대지 않은 채 저장된 경우에만 바꿔 줍니다 */
  var V1_TYPES = { '이종 불량': '이종/누락', '용접 균열': '용접불량', '용접 불량': '용접불량', '가공 치수불량': '위치/치수', '도장 불량': '도장/외관불량',
    '외관 불량': '도장/외관불량', '조립·체결 불량': '미조임', '누유·누수': '누유/누수', '녹·부식': '녹/발청' };
  /* 2차(v2) 기본 사전 8종 · 담당부서 → 3차(v3): 도장/외관불량 · 녹/발청을 기타 앞에 더하고, 담당부서를 협력사 내부 부서로 */
  var V2_NAMES = ['위치/치수', '간섭/정렬', '이종/누락', '용접불량', '누유/누수', '미조임', '성능이상', '기타'];
  var V2_DEPTS = ['부품품질팀', '협력사', '구매팀', '생산기술팀', '설계팀'];
  var V2_DEPT_MAP = { '부품품질팀': '품질', '생산기술팀': '생산기술', '설계팀': '설계', '구매팀': '구매' };
  var ACTION_DEPT = {}; DEFAULT_TYPES.forEach(function (t) { t.actions.forEach(function (a) { ACTION_DEPT[a[0]] = a[1]; }); });
  function v3Dept(d, text) { return V2_DEPT_MAP[d] || (d === '협력사' ? ACTION_DEPT[text] || '품질' : d); }
  function isV1Defaults(types) {
    var names = (types || []).map(function (t) { return t && str(t.name); });
    return names.length === 9 && names.every(function (n) { return V1_TYPES[n]; });
  }
  function emptyDb() {
    return { schemaVersion: SCHEMA_VERSION, records: [], seq: 0, types: defaultTypes(), depts: DEFAULT_DEPTS.slice(), rules: cleanRules({}),
      plans: {}, aliases: { suppliers: {}, parts: {} }, mask: true, mapping: null, files: [], updates: [], replyBy: '', aiTarget: 'copilot', mail: defaultMail() };
  }
  function cleanType(t) {
    if (!t || typeof t !== 'object' || !str(t.name) || str(t.name) === UNCLASSIFIED) return null;
    return { name: str(t.name), dept: str(t.dept) || '품질', keywords: Array.isArray(t.keywords) ? uniq(t.keywords.map(str).filter(Boolean)) : [],
      actions: Array.isArray(t.actions) ? t.actions.map(function (a) { return a && str(a.text) ? { text: str(a.text), dept: str(a.dept) || '품질', start: Math.max(0, parseInt(a.start, 10) || 0), weeks: Math.max(1, parseInt(a.weeks, 10) || 1) } : null; }).filter(Boolean) : [] };
  }
  function restoreDb(o) {
    var db = emptyDb();
    if (!o || typeof o !== 'object') return db;
    if (Array.isArray(o.records)) db.records = o.records.filter(function (r) { return r && typeof r === 'object'; }).map(function (r) {
      var c = { id: str(r.id), date: parseDate(r.date), partName: str(r.partName), partNo: str(r.partNo), problem: str(r.problem), supplier: str(r.supplier) };
      if (typeof r.qty === 'number') c.qty = r.qty;
      if (str(r.typeManual)) c.typeManual = str(r.typeManual);
      if (str(r.typeIn)) c.typeIn = str(r.typeIn);
      if (str(r.action)) c.action = str(r.action);
      if (str(r.srcNo)) c.srcNo = str(r.srcNo);
      if (normCampus(r.campus)) c.campus = normCampus(r.campus);
      if (r.src === 'raw') c.src = 'raw';
      return c;
    });
    if (typeof o.seq === 'number') db.seq = o.seq;
    db.records.forEach(function (r) { if (!r.id) r.id = 'd' + (++db.seq); });
    var v1 = Array.isArray(o.types) && !(o.schemaVersion >= 2) && isV1Defaults(o.types);
    var old3 = !(o.schemaVersion >= 3);
    if (Array.isArray(o.types) && !v1) { var ts = o.types.map(cleanType).filter(Boolean); if (ts.length) db.types = ts; }
    if (v1) db.records.forEach(function (r) { if (r.typeManual && V1_TYPES[r.typeManual]) r.typeManual = V1_TYPES[r.typeManual]; });
    // v2 → v3: 기본 8종을 그대로 쓰고 있었으면 도장/외관불량 · 녹/발청을 기타 앞에 넣고, 그 키워드는 기타에서 뺍니다
    if (old3 && !v1 && typeNames(db.types).join('|') === V2_NAMES.join('|')) {
      var add = defaultTypes().filter(function (t) { return t.name === '도장/외관불량' || t.name === '녹/발청'; });
      var moved = {}; add.forEach(function (t) { t.keywords.forEach(function (k) { moved[normText(k)] = 1; }); });
      var etc = db.types[db.types.length - 1];
      etc.keywords = etc.keywords.filter(function (k) { return !moved[normText(k)] && !/^녹/.test(k); });
      db.types.splice.apply(db.types, [db.types.length - 1, 0].concat(add));
    }
    var deptsDefault = !Array.isArray(o.depts) || (o.depts.length === V2_DEPTS.length && V2_DEPTS.every(function (d) { return o.depts.indexOf(d) >= 0; }));
    if (Array.isArray(o.depts) && !(old3 && deptsDefault)) { var ds = uniq(o.depts.map(str).filter(Boolean)); if (ds.length) db.depts = ds; }
    db.rules = old3 ? cleanRules({}) : cleanRules(o.rules);
    if (o.plans && typeof o.plans === 'object') Object.keys(o.plans).forEach(function (s) {
      var p = o.plans[s]; if (!p || typeof p !== 'object') return;
      db.plans[s] = { baseDate: parseDate(p.baseDate), from: parseDate(p.from), to: parseDate(p.to), rows: Array.isArray(p.rows) ? p.rows.map(cleanPlanRow) : [],
        confirmed: Array.isArray(p.confirmed) ? p.confirmed.filter(function (c) { return c && Array.isArray(c.rows); }).map(function (c, i) { return { version: i + 1, at: str(c.at), rows: c.rows.map(cleanPlanRow) }; }) : [],
        aiSummary: str(p.aiSummary), aiChecks: Array.isArray(p.aiChecks) ? p.aiChecks.map(str).filter(Boolean) : [], aiAnswer: p.aiAnswer == null ? '' : String(p.aiAnswer) };
    });
    if (v1) Object.keys(db.plans).forEach(function (s) {
      var p = db.plans[s];
      [p.rows].concat(p.confirmed.map(function (c) { return c.rows; })).forEach(function (rs) { rs.forEach(function (r) { if (V1_TYPES[r.type]) r.type = V1_TYPES[r.type]; }); });
    });
    // v2 → v3: 옛 기본 담당부서(부품품질팀·협력사·…)를 협력사 내부 부서(품질·생산기술·영업·설계·구매)로 — 확정본도 같이 바꿔 「수정 있음」이 생기지 않게
    if (old3 && deptsDefault) {
      db.types.forEach(function (t) { t.dept = v3Dept(t.dept, ''); t.actions.forEach(function (a) { a.dept = v3Dept(a.dept, a.text); }); });
      Object.keys(db.plans).forEach(function (s) {
        var p = db.plans[s];
        [p.rows].concat(p.confirmed.map(function (c) { return c.rows; })).forEach(function (rs) { rs.forEach(function (r) { r.dept = v3Dept(r.dept, r.action); }); });
      });
    }
    if (o.aliases && typeof o.aliases === 'object') db.aliases = { suppliers: assignAliases(o.aliases.suppliers, [], ALIAS.suppliers), parts: assignAliases(o.aliases.parts, [], ALIAS.parts) };
    db.mask = o.mask !== false;
    if (o.mapping && typeof o.mapping === 'object') db.mapping = o.mapping;
    if (Array.isArray(o.files)) db.files = o.files.slice(-50);
    if (Array.isArray(o.updates)) db.updates = o.updates.filter(function (u) { return u && typeof u === 'object' && str(u.at); }).slice(-200);
    db.replyBy = str(o.replyBy);
    if (o.aiTarget && AI_TARGETS.some(function (t) { return t.id === o.aiTarget; })) db.aiTarget = o.aiTarget;
    db.mail = cleanMail(o.mail);
    if (o._sample) db._sample = true;
    return db;
  }
  /* ── 12. 아웃룩 메일 초안(.eml) — 4차(2026-10-27) ─────────────────────
     브라우저는 아웃룩을 직접 열거나 첨부를 넣을 수 없습니다. 그래서 받는 사람 · 참조 · 제목 · 본문 · MASTER PLAN 엑셀을 담은
     메일 파일(.eml, RFC 5322 · MIME multipart/mixed)을 만들어 내려받게 합니다. 「X-Unsent: 1」 머리글이 있어 아웃룩이
     「보낸 메일」이 아니라 고칠 수 있는 새 메일(초안)로 엽니다. 첨부 없이 쓰는 mailto: 링크도 함께 만듭니다.
     설정 = 기본 수신 1~3 · 참조 1~3 (이름 · EMAIL) + 협력사마다 바꿀 수신 · 참조(선택) + 제목 · 본문(자리표시 {협력사} 등) */
  var MAIL_SLOTS = 3;
  var MAIL_VARS = [['협력사', '협력사 이름'], ['마감', '회신 마감일(메일 만들 때 고름)'], ['기준일', '오늘 날짜'], ['과제수', '그 협력사의 개선과제 수'],
    ['미완료', '마감이 지난 미완료 과제 수'], ['진행중', '진행중 과제 수'], ['파일명', '첨부 엑셀 파일 이름'], ['보내는사람', '설정의 보내는 사람 이름']];
  var DEFAULT_MAIL = {
    subject: '[공정불량 개선] {협력사} 개선 MASTER PLAN 회신 요청 ({마감}까지)',
    body: '{협력사} 품질 담당자님, 안녕하세요.\n\n' +
      '{기준일} 기준 공정불량 분석 결과와 개선 MASTER PLAN 초안(개선과제 {과제수}건)을 첨부합니다.\n' +
      '첨부 파일 「{파일명}」의 MASTER PLAN 시트에서 노란 칸(개선대책 · 담당부서 · 일정 · 실적 · 진행상태)을 적어 {마감}까지 회신 부탁드립니다.\n' +
      '숨긴 ID 칸과 기준 칸은 지우지 말아 주세요. 새 활동은 표 아래 빈 줄에 적으면 됩니다.\n\n감사합니다.\n{보내는사람} 드림'
  };
  function cleanPeople(list) {
    var out = [];
    for (var i = 0; i < MAIL_SLOTS; i++) { var p = (list || [])[i] || {}; out.push({ name: str(p.name), email: str(p.email) }); }
    return out;
  }
  function emptyPeople() { return cleanPeople([]); }
  function defaultMail() { return { from: { name: '', email: '' }, to: emptyPeople(), cc: emptyPeople(), subject: DEFAULT_MAIL.subject, body: DEFAULT_MAIL.body, per: {} }; }
  function cleanMail(o) {
    var m = defaultMail();
    if (!o || typeof o !== 'object') return m;
    if (o.from && typeof o.from === 'object') m.from = { name: str(o.from.name), email: str(o.from.email) };
    m.to = cleanPeople(o.to); m.cc = cleanPeople(o.cc);
    if (typeof o.subject === 'string') m.subject = o.subject;
    if (typeof o.body === 'string') m.body = o.body.replace(/\r\n?/g, '\n');
    if (o.per && typeof o.per === 'object') Object.keys(o.per).forEach(function (s) {
      var p = o.per[s]; if (!str(s) || !p || typeof p !== 'object') return;
      var x = { to: cleanPeople(p.to), cc: cleanPeople(p.cc) };
      if (filled(x.to).length || filled(x.cc).length) m.per[str(s)] = x;
    });
    return m;
  }
  function filled(list) { return (list || []).filter(function (p) { return str(p.email); }); }
  var EMAIL_RE = /^[^\s@<>(),;:"\[\]\\]+@[^\s@<>(),;:"\[\]\\]+\.[^\s@<>(),;:"\[\]\\]+$/;
  function isEmail(s) { return EMAIL_RE.test(str(s)); }
  /* 그 협력사에 쓸 수신 · 참조 — 협력사 칸에 수신을 하나라도 적었으면 수신은 그것, 참조도 따로(비었으면 기본 참조) */
  function mailFor(mail, supplier) {
    mail = mail || defaultMail();
    var p = mail.per && mail.per[supplier];
    var to = p && filled(p.to).length ? filled(p.to) : filled(mail.to);
    var cc = p && filled(p.cc).length ? filled(p.cc) : filled(mail.cc);
    return { to: to, cc: cc, override: !!p };
  }
  function checkMail(mail, supplier) {
    var r = mailFor(mail, supplier), e = [];
    if (!r.to.length) e.push('받는 사람 EMAIL 이 없습니다 — 설정 「메일 받는 사람」에 수신 1을 적어 주세요.');
    r.to.concat(r.cc).forEach(function (p) { if (!isEmail(p.email)) e.push('EMAIL 모양이 아닙니다: ' + p.email); });
    if (str(mail.from && mail.from.email) && !isEmail(mail.from.email)) e.push('보내는 사람 EMAIL 모양이 아닙니다: ' + mail.from.email);
    return e;
  }
  /* 자리표시 {협력사} 등을 값으로. 모르는 {…} 는 그대로 둡니다 */
  function fillTemplate(tpl, vars) {
    return String(tpl == null ? '' : tpl).replace(/\{([^{}\s]+)\}/g, function (m, k) { return Object.prototype.hasOwnProperty.call(vars, k) ? String(vars[k]) : m; });
  }
  function mailVars(o) {
    var st = planStats(o.rows || [], o.today);
    return { '협력사': o.supplier, '마감': o.due || '', '기준일': o.today, '과제수': st.total, '미완료': st.late, '진행중': st.doing, '파일명': o.fileName || '', '보내는사람': str(o.fromName) };
  }
  /* UTF-8 바이트 · base64 (node · 브라우저 모두, 외부 라이브러리 없이) */
  function utf8Bytes(s) {
    s = String(s); var out = [];
    for (var i = 0; i < s.length; i++) {
      var c = s.charCodeAt(i);
      if (c >= 0xd800 && c < 0xdc00 && i + 1 < s.length) { var d = s.charCodeAt(i + 1); if (d >= 0xdc00 && d < 0xe000) { c = 0x10000 + ((c - 0xd800) << 10) + (d - 0xdc00); i++; } }
      if (c < 0x80) out.push(c);
      else if (c < 0x800) out.push(0xc0 | c >> 6, 0x80 | c & 63);
      else if (c < 0x10000) out.push(0xe0 | c >> 12, 0x80 | c >> 6 & 63, 0x80 | c & 63);
      else out.push(0xf0 | c >> 18, 0x80 | c >> 12 & 63, 0x80 | c >> 6 & 63, 0x80 | c & 63);
    }
    return out;
  }
  var B64 = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
  function base64(bytes) {
    var s = '', i, n = bytes.length;
    for (i = 0; i + 2 < n; i += 3) { var v = bytes[i] << 16 | bytes[i + 1] << 8 | bytes[i + 2]; s += B64[v >> 18 & 63] + B64[v >> 12 & 63] + B64[v >> 6 & 63] + B64[v & 63]; }
    if (n - i === 1) { var a = bytes[i] << 16; s += B64[a >> 18 & 63] + B64[a >> 12 & 63] + '=='; }
    else if (n - i === 2) { var b = bytes[i] << 16 | bytes[i + 1] << 8; s += B64[b >> 18 & 63] + B64[b >> 12 & 63] + B64[b >> 6 & 63] + '='; }
    return s;
  }
  function wrap76(s) { var out = []; for (var i = 0; i < s.length; i += 76) out.push(s.slice(i, i + 76)); return out.join('\r\n'); }
  /* RFC 2047 encoded-word(B) — 글자를 반으로 자르지 않게 글자 단위로 묶고, 한 낱말은 75자 이하 */
  function encodeWord(s) {
    s = String(s);
    if (/^[\x20-\x7e]*$/.test(s)) return s;
    var words = [], cur = [], chars = Array.from ? Array.from(s) : s.split('');
    chars.forEach(function (ch) {
      var b = utf8Bytes(ch);
      if (Math.ceil((cur.length + b.length) / 3) * 4 > 45) { words.push(cur); cur = []; } // 45 + 12(=?UTF-8?B??=) ≤ 75
      cur = cur.concat(b);
    });
    if (cur.length) words.push(cur);
    return words.map(function (w) { return '=?UTF-8?B?' + base64(w) + '?='; }).join('\r\n ');
  }
  function addrList(list) {
    return list.map(function (p) {
      var n = str(p.name), e = str(p.email);
      if (!n) return e;
      return (/^[\x20-\x7e]*$/.test(n) ? '"' + n.replace(/(["\\])/g, '\\$1') + '"' : encodeWord(n)) + ' <' + e + '>';
    }).join(',\r\n ');
  }
  function mailDate(d) {
    var D = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'], M = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
    function p(n) { return (n < 10 ? '0' : '') + n; }
    var off = -d.getTimezoneOffset(), sign = off >= 0 ? '+' : '-'; off = Math.abs(off);
    return D[d.getDay()] + ', ' + p(d.getDate()) + ' ' + M[d.getMonth()] + ' ' + d.getFullYear() + ' ' + p(d.getHours()) + ':' + p(d.getMinutes()) + ':' + p(d.getSeconds()) + ' ' + sign + p(Math.floor(off / 60)) + p(off % 60);
  }
  /* o = { from:{name,email}?, to:[{name,email}], cc:[…], subject, body, attachments:[{name, type, bytes(Uint8Array)}], date:Date?, boundary? } → 메일 글(CRLF) */
  function buildEml(o) {
    var bnd = o.boundary || ('----=_data10-07_' + Date.now().toString(36) + Math.random().toString(36).slice(2, 10));
    var H = [];
    if (o.from && str(o.from.email)) H.push('From: ' + addrList([o.from]));
    if (o.to && o.to.length) H.push('To: ' + addrList(o.to));
    if (o.cc && o.cc.length) H.push('Cc: ' + addrList(o.cc));
    H.push('Subject: ' + encodeWord(o.subject || ''));
    H.push('Date: ' + mailDate(o.date || new Date()));
    H.push('MIME-Version: 1.0');
    H.push('X-Unsent: 1');
    H.push('Content-Type: multipart/mixed; boundary="' + bnd + '"');
    var parts = ['--' + bnd, 'Content-Type: text/plain; charset="UTF-8"', 'Content-Transfer-Encoding: base64', '', wrap76(base64(utf8Bytes(String(o.body || '').replace(/\r\n?/g, '\n').replace(/\n/g, '\r\n'))))];
    (o.attachments || []).forEach(function (a) {
      var fn = encodeURIComponent(a.name).replace(/['()*]/g, function (c) { return '%' + c.charCodeAt(0).toString(16).toUpperCase(); });
      var nm = encodeWord(a.name).replace(/\r\n /g, ' ');
      parts.push('--' + bnd,
        'Content-Type: ' + (a.type || 'application/octet-stream') + '; name="' + nm + '"',
        'Content-Transfer-Encoding: base64',
        'Content-Disposition: attachment; filename*=UTF-8\'\'' + fn,
        '', wrap76(base64(a.bytes)));
    });
    parts.push('--' + bnd + '--', '');
    return H.join('\r\n') + '\r\n\r\n' + 'This is a multi-part message in MIME format.\r\n\r\n' + parts.join('\r\n');
  }
  /* 첨부 없는 mailto: — 받는 사람 · 참조 · 제목 · 본문만. 아웃룩 등 기본 메일 프로그램이 열립니다(본문이 길면 잘릴 수 있음) */
  function mailtoUrl(o) {
    function enc(s) { return encodeURIComponent(s).replace(/%20/g, '%20'); }
    var q = [];
    if (o.cc && o.cc.length) q.push('cc=' + o.cc.map(function (p) { return enc(p.email); }).join(','));
    q.push('subject=' + enc(o.subject || ''));
    q.push('body=' + enc(String(o.body || '').replace(/\r?\n/g, '\r\n')));
    return 'mailto:' + (o.to || []).map(function (p) { return enc(p.email); }).join(',') + '?' + q.join('&');
  }
  function safeFileName(s) { return str(s).replace(/[\\/:*?"<>|\s]+/g, '').slice(0, 60) || '협력사'; }

  function suppliersOf(records) { return uniq(records.map(function (r) { return r.supplier; }).filter(Boolean)).sort(); }

  return {
    VERSION: VERSION,
    // 날짜
    parseDate: parseDate, toDateStr: toDateStr, addDays: addDays, daysBetween: daysBetween, mondayOf: mondayOf, monthOf: monthOf, monthAdd: monthAdd, weekInfo: weekInfo,
    // 읽기
    FIELDS: FIELDS, PASTE_ORDER: PASTE_ORDER, fieldLabel: fieldLabel, detectColumns: detectColumns, parseSheet: parseSheet, pickSheet: pickSheet, recordKey: recordKey, mergeRecords: mergeRecords, normHead: normHead,
    CAMPUSES: CAMPUSES, campusLabel: campusLabel, normCampus: normCampus, replaceCampus: replaceCampus, campusCounts: campusCounts, campusMatch: campusMatch,
    splitTsv: splitTsv, parsePaste: parsePaste, checkRecordInput: checkRecordInput, recordFromInput: recordFromInput,
    // 분류
    UNCLASSIFIED: UNCLASSIFIED, DEFAULT_DEPTS: DEFAULT_DEPTS, defaultTypes: defaultTypes, classify: classify, typeOf: typeOf, enrich: enrich, typeNames: typeNames, findTypeName: findTypeName,
    parseActionLines: parseActionLines, formatActionLines: formatActionLines, parseKeywords: parseKeywords,
    // 분석
    DEFAULT_RULES: DEFAULT_RULES, cleanRules: cleanRules, analyze: analyze, findIssues: findIssues, issueTitle: issueTitle, issueText: issueText, dateBounds: dateBounds, suppliersOf: suppliersOf,
    // 계획
    STATUS: STATUS, PHASE: PHASE, PHASE_ORDER: PHASE_ORDER, phaseOf: phaseOf, notStarted: notStarted, PHASE_IN: PHASE_IN,
    buildDraft: buildDraft, makeRow: makeRow, cleanPlanRow: cleanPlanRow, checkPlanRow: checkPlanRow, planStats: planStats, isOverdue: isOverdue,
    ganttColumns: ganttColumns, span: span, groupSpans: groupSpans,
    confirmPlan: confirmPlan, lastConfirmed: lastConfirmed, planChangedSince: planChangedSince, planStateText: planStateText,
    // AI
    ALIAS: ALIAS, AI_TARGETS: AI_TARGETS, assignAliases: assignAliases, updateAliases: updateAliases, maskText: maskText, unmaskText: unmaskText, findLeaks: findLeaks,
    buildAiPromptInfo: buildAiPromptInfo, parseAiPlan: parseAiPlan, matchType: matchType, matchDept: matchDept,
    // Excel 추출·진행사항 업로드·저장
    reportSheets: reportSheets, reportBook: reportBook, backupSheet: backupSheet, REPORT_TITLE: REPORT_TITLE, SHEET_LIST: SHEET_LIST, SHEET_PLAN: SHEET_PLAN, PCOLS: PCOLS, sourceText: sourceText,
    parseProgressSheet: parseProgressSheet, diffProgress: diffProgress, applyProgress: applyProgress, fingerprint: fingerprint,
    emptyDb: emptyDb, restoreDb: restoreDb, SCHEMA_VERSION: SCHEMA_VERSION,
    // 메일(4차)
    MAIL_SLOTS: MAIL_SLOTS, MAIL_VARS: MAIL_VARS, DEFAULT_MAIL: DEFAULT_MAIL, defaultMail: defaultMail, cleanMail: cleanMail, mailFor: mailFor, checkMail: checkMail, isEmail: isEmail,
    fillTemplate: fillTemplate, mailVars: mailVars, buildEml: buildEml, mailtoUrl: mailtoUrl, encodeWord: encodeWord, base64: base64, utf8Bytes: utf8Bytes, safeFileName: safeFileName
  };
});
