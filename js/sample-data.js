/* 예시 데이터 — 전부 지어낸 가상 값입니다(협력사 A~F, 품번 TP-…). 실제 회사·협력사 자료가 아닙니다.
   고정 씨앗 난수라 늘 같은 데이터가 나옵니다. 분석 규칙이 제대로 잡는지 보려고 일부러 심어 둔 것:
   - 협력사 C : 용접불량(균열)이 8~10월에 늘어남 → WORST 1, 같은 품번 TP-3301 반복(동일품번)
   - 협력사 A : 이종/누락 WORST 1, 같은 품번 TP-1102 반복
   - 협력사 D : 위치/치수 WORST 1, 같은 품번 TP-4405 반복
   - 협력사 F : 건수가 적음(3건) — 그래도 WORST 1·2 는 잡힘
   - 어느 유형에도 맞지 않는 내용 몇 건(미분류)
   불량유형은 수강생이 정한 10종(위치/치수 · 간섭/정렬 · 이종/누락 · 용접불량 · 누유/누수 · 미조임 · 성능이상 · 도장/외관불량 · 녹/발청 · 기타)입니다. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.SQSample = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';
  function rng(seed) {
    return function () {
      seed |= 0; seed = seed + 0x6D2B79F5 | 0;
      var t = Math.imul(seed ^ seed >>> 15, 1 | seed);
      t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t;
      return ((t ^ t >>> 14) >>> 0) / 4294967296;
    };
  }
  var TEXT = {
    '위치/치수': ['홀 위치 치수 공차 초과', '가공 치수 규격 미달(외경)', '탭 나사산 불량', '평면도 불량', '브래킷 길이 치수 초과'],
    '간섭/정렬': ['조립 시 상대 부품과 간섭', '브래킷 정렬 틀어짐', '오조립(방향 반대)', '체결면 단차 발생'],
    '이종/누락': ['이종품 혼입 (타 품번 2개 섞임)', '오품 납품 — 좌우 품번 바뀜', '와셔 누락', '라벨 오류로 품번 상이', '부품 미장착 상태로 입고'],
    '용접불량': ['용접 기공 다수', '언더컷 발생', '스패터 과다 부착', '미용접 구간 있음', '각장 부족'],
    '누유/누수': ['유압 파이프 연결부 누유', '기밀 시험 리크 발생', '피팅부 오일 샘'],
    '미조임': ['볼트 체결 토크 부족', '너트 풀림', '호스 클램프 미조임'],
    '성능이상': ['밸브 작동 불량', '작동 시 소음 발생', '압력 저하로 성능 미달'],
    '도장/외관불량': ['도장 흐름 발생', '표면 찍힘', '스크래치 다수', '도막 두께 부족'],
    '녹/발청': ['보관 중 녹 발생', '표면 발청', '나사부 부식'],
    '기타': ['포장 상태 불량', '이물 부착', '운송 중 변형']
  };
  var CRACK = ['용접부 균열 발생', '용접 비드 끝단 크랙', '브래킷 용접부 갈라짐', '용접 후 균열(열영향부)'];
  var ACTIONS = ['', '', '전수 선별', '재작업 후 재입고', '협력사 반품', '특채 사용(현장 승인)', '임시 보수'];
  var OTHER = ['검사성적서 미첨부', '원인 확인 중', '납품 서류 확인 필요'];
  var SUP = {
    '협력사 A': { n: 26, parts: [['TP-1101', '브래킷'], ['TP-1102', '브래킷 ASSY'], ['TP-1103', '커버'], ['TP-1104', '핀']], w: { '이종/누락': 5, '녹/발청': 1, '기타': 1, '위치/치수': 1, '미조임': 1 } },
    '협력사 B': { n: 30, parts: [['TP-2201', '탱크'], ['TP-2202', '커버'], ['TP-2203', '도어 판넬']], w: { '도장/외관불량': 3, '녹/발청': 2, '간섭/정렬': 3, '성능이상': 1 } },
    '협력사 C': { n: 22, parts: [['TP-3301', '프레임'], ['TP-3302', '브래킷'], ['TP-3303', '서포트']], w: { '용접불량': 2, '위치/치수': 2, '간섭/정렬': 1 } },
    '협력사 D': { n: 28, parts: [['TP-4401', '샤프트'], ['TP-4402', '부싱'], ['TP-4403', '하우징'], ['TP-4405', '플랜지']], w: { '위치/치수': 5, '간섭/정렬': 1, '도장/외관불량': 1 } },
    '협력사 E': { n: 24, parts: [['TP-5501', '유압 파이프'], ['TP-5502', '호스 ASSY'], ['TP-5503', '밸브 블록']], w: { '미조임': 3, '누유/누수': 3, '성능이상': 1 } },
    '협력사 F': { n: 3, parts: [['TP-6601', '가이드'], ['TP-6602', '스페이서']], w: { '도장/외관불량': 1, '위치/치수': 1 } }
  };
  var START = '2026-01-05', END = '2026-10-16';
  function addDays(s, n) { var d = new Date(Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10)) + n * 86400000); return d.toISOString().slice(0, 10); }
  function span() { return Math.round((Date.UTC(2026, 9, 16) - Date.UTC(2026, 0, 5)) / 86400000); }
  function pick(r, arr) { return arr[Math.floor(r() * arr.length)]; }
  function wpick(r, w) {
    var ks = Object.keys(w), tot = ks.reduce(function (s, k) { return s + w[k]; }, 0), x = r() * tot;
    for (var i = 0; i < ks.length; i++) { x -= w[ks[i]]; if (x < 0) return ks[i]; }
    return ks[ks.length - 1];
  }
  /* 품번 풀 — 실제 협력사는 품번이 많아 같은 품번이 겹치는 일이 드뭅니다. 앞의 대표 품번 + 같은 계열 품번 60개(TP-x1NN) */
  function pool(parts) {
    var head = parts[0][0].slice(0, 4), names = parts.map(function (q) { return q[1]; }), out = parts.slice();
    for (var i = 0; i < 60; i++) out.push([head + '1' + (i < 10 ? '0' : '') + i, names[i % names.length]]);
    return out;
  }
  /* 발생일자 순 레코드 [{date, supplier, partNo, partName, problem, qty}] */
  function records() {
    var r = rng(20261020), out = [], days = span();
    Object.keys(SUP).forEach(function (s) {
      var p = SUP[s];
      for (var i = 0; i < p.n; i++) {
        var t = wpick(r, p.w), part = pick(r, pool(p.parts));
        out.push({ date: addDays(START, Math.floor(r() * days)), supplier: s, partNo: part[0], partName: part[1], problem: pick(r, TEXT[t]), qty: 1 + Math.floor(r() * 6), action: pick(r, ACTIONS) });
      }
    });
    // 심어 둔 신호
    function put(date, s, pn, name, problem, qty, action) { out.push({ date: date, supplier: s, partNo: pn, partName: name, problem: problem, qty: qty, action: action || '' }); }
    put('2026-03-11', '협력사 C', 'TP-3302', '브래킷', CRACK[0], 2, '재작업 후 재입고');
    ['2026-08-04', '2026-08-21', '2026-09-02', '2026-09-17', '2026-09-29', '2026-10-08', '2026-10-14'].forEach(function (d, i) {
      put(d, '협력사 C', i % 3 === 2 ? 'TP-3303' : 'TP-3301', i % 3 === 2 ? '서포트' : '프레임', CRACK[i % 4], 1 + (i % 3), i % 2 ? '전수 선별' : '');
    });
    ['2026-02-17', '2026-05-12', '2026-07-08', '2026-09-23'].forEach(function (d, i) { put(d, '협력사 A', 'TP-1102', '브래킷 ASSY', TEXT['이종/누락'][i % 3], 4 + i, '협력사 반품'); });
    ['2026-04-06', '2026-06-15', '2026-08-26'].forEach(function (d, i) { put(d, '협력사 D', 'TP-4405', '플랜지', TEXT['위치/치수'][i], 3); });
    put('2026-05-20', '협력사 B', 'TP-2202', '커버', OTHER[0], 1);
    put('2026-07-14', '협력사 E', 'TP-5502', '호스 ASSY', OTHER[1], 10);
    put('2026-09-09', '협력사 F', 'TP-6601', '가이드', OTHER[2], 1);
    out.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : a.supplier < b.supplier ? -1 : 1; });
    return out;
  }
  /* 엑셀 모양 — 제목 줄과 빈 줄이 있고, 머리행 이름을 일부러 표준과 다르게(업체명·품 번·불량 내용) 적었습니다 */
  var HEAD = ['No', '발생일', '업체명', '품 번', '품명', '불량 내용', '수량', '조치 내용'];
  function aoa() {
    var rows = [['■ 협력사 공정불량 현황 (가상 예시 — 실제 회사 자료 아님)'], [], HEAD];
    records().forEach(function (x, i) { rows.push([i + 1, x.date, x.supplier, x.partNo, x.partName, x.problem, x.qty, x.action || null]); });
    return rows;
  }
  /* 4차(2026-10-27) 캠퍼스별 RAW DATA 모양 — 회사는 통합됐지만 시스템이 통합 전이라 울산 · 인천의 열 구성이 다릅니다.
     열 이름은 수강생이 보내 준 가상 파일의 모양을 따랐고, 값은 위 예시 데이터(협력사 A~F · 품번 TP-…)를 셋 중 하나는 인천, 나머지는 울산으로 나눈 것입니다.
     일부러 심은 것: 울산 파일 끝에 같은 줄 한 번 더(→ 캠퍼스 안 중복이라 1건 건너뜀),
                     울산의 첫 줄과 같은 내용을 인천에도 한 줄(→ 캠퍼스가 다르니 둘 다 남음). */
  var U_HEAD = ['저장 상태', '문제점 번호', '발생 일자', '월', '생산라인', '품번', '부품 상세', '협력회사', '협력회사명', '조치 담당자명', '문제점 상세', '수량', '조치 방안', '조치방안 상세', '부적합 유형', '귀책업체명', '플랜트'];
  var I_HEAD = ['Status', 'Issued Date', 'QM noti No.', 'Plant', 'Vendor Code', 'Vendor Name', 'Defect Dev.', 'Defect Dev. Desc', 'Inspector', 'Inspector Name', 'Supervisor', 'Material',
    'Material Description', 'Qty', 'Problem Description', 'Attach Document', 'Iimprovement Doc.', 'Imp.Doc.Date', 'Cause Analysis', 'Corrective Action', 'Supv.App.Date', 'Description'];
  function campusOf(i) { return i % 3 === 2 ? '인천' : '울산'; }
  function vcode(s) { return 'V-10' + s.slice(-1); }
  function uRow(x, n) { return ['S', 'U-26-' + ('000' + n).slice(-4), x.date, +x.date.slice(5, 7), 'L' + (1 + n % 3), x.partNo, x.partName, vcode(x.supplier), x.supplier, '품질 담당', x.problem, x.qty, x.action ? 'A' : '', x.action || '', 'N' + (n % 4), x.supplier, 'U100']; }
  function iRow(x, n) { return ['A', x.date, 'QN-' + (7000 + n), 'I200', vcode(x.supplier), x.supplier, 'D' + (n % 3), '', 'IN-' + (n % 5), 'Inspector ' + (n % 5), 'Supervisor ' + (n % 2), x.partNo, x.partName, x.qty, x.problem, '', '', '', '', x.action || '', '', '']; }
  function campusRecords(campus) { return records().filter(function (x, i) { return campusOf(i) === campus; }); }
  function campusAoa(campus) {
    var list = campusRecords(campus), first = campusRecords('울산')[0];
    if (campus === '인천') {
      var rows = [I_HEAD];
      list.forEach(function (x, i) { rows.push(iRow(x, i + 1)); });
      rows.push(iRow(first, list.length + 1)); // 울산 첫 줄과 같은 내용(관리번호만 인천 것)
      return rows;
    }
    var u = [['가상 데이터 예시 — 울산캠퍼스 공정불량 RAW (협력사 A~F · 품번 TP-…, 실제 회사 자료 아님)'], [], U_HEAD];
    list.forEach(function (x, i) { u.push(uRow(x, i + 1)); });
    u.push(uRow(list[0], 1)); // 같은 줄 한 번 더 — 캠퍼스 안 중복
    return u;
  }
  return { records: records, aoa: aoa, HEAD: HEAD, START: START, END: END, campusOf: campusOf, campusAoa: campusAoa, campusRecords: campusRecords, U_HEAD: U_HEAD, I_HEAD: I_HEAD };
});
