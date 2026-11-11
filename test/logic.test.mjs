// 실행: node test/logic.test.mjs   (의존성 없음 — vendor 의 SheetJS 만 씀)
// 실제 공정불량 Raw data 는 아직 받지 않았습니다. 여기서는 가상 표와 예시 데이터(협력사 A~F)로 검사합니다.
import assert from 'node:assert/strict';
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const L = require('../js/logic.js');
const Sample = require('../js/sample-data.js');
const E = require('../js/ai-endpoint.js');
const XLSX = require('../vendor/xlsx.full.min.js');
const XO = require('../js/xlsx-out.js');

let passed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log('  ok  ' + name); }
  catch (e) { console.error('  FAIL ' + name + '\n       ' + e.message); process.exitCode = 1; }
}
let idn = 0;
const newId = () => 't' + (++idn);
const TYPES = L.defaultTypes();

console.log('날짜·주차');
test('여러 날짜 표기 · 엑셀 일련번호 · 없는 날짜는 빈 값', () => {
  assert.equal(L.parseDate('2026-10-05'), '2026-10-05');
  assert.equal(L.parseDate('2026.10.5'), '2026-10-05');
  assert.equal(L.parseDate('2026/10/05 13:20'), '2026-10-05');
  assert.equal(L.parseDate('2026년 10월 5일'), '2026-10-05');
  assert.equal(L.parseDate('26.10.05'), '2026-10-05');
  assert.equal(L.parseDate('20261005'), '2026-10-05');
  assert.equal(L.parseDate(46300), '2026-10-05');
  assert.equal(L.parseDate('2026-02-30'), '');
  assert.equal(L.parseDate('확인 중'), '');
});
test('주차 — 목요일이 든 달 기준 (시안의 「2026년 10월 5주차」)', () => {
  assert.equal(L.weekInfo('2026-10-26').label, '2026년 10월 5주차');
  assert.equal(L.weekInfo('2026-11-01').label, '2026년 10월 5주차'); // 일요일
  assert.equal(L.weekInfo('2026-09-28').label, '2026년 10월 1주차'); // 10/1(목)이 든 주
  assert.equal(L.weekInfo('2026-11-02').label, '2026년 11월 1주차');
  assert.equal(L.weekInfo('2027-01-01').label, '2026년 12월 5주차'); // 12/31(목)이 든 주
  assert.equal(L.mondayOf('2026-10-25'), '2026-10-19');
});

console.log('엑셀 열 찾기');
test('제목 줄 아래 머리행 · 다른 이름(업체명·품 번·불량 내용)도 연결', () => {
  const d = L.detectColumns(Sample.aoa());
  assert.equal(d.headerRow, 2);
  assert.deepEqual(d.missing, []);
  assert.equal(d.map.supplier, 2); assert.equal(d.map.partNo, 3); assert.equal(d.map.problem, 5); assert.equal(d.map.qty, 6); assert.equal(d.map.action, 7);
});
test('불량유형 칸은 「불량유형」으로(내용과 따로) · 조치 칸 · 품번 칸은 없어도 읽음', () => {
  const aoa = [['발생일자', '협력사', '불량유형', '불량내용', '조치사항'], ['2026-10-05', '협력사 A', '용접 불량', '비드 끝 크랙', '재작업'], ['2026-10-06', '협력사 A', '미조임', '', '']];
  const p = L.parseSheet(aoa);
  assert.deepEqual(p.det.missing, []);
  assert.equal(p.det.map.type, 2); assert.equal(p.det.map.problem, 3);
  assert.equal(p.records[0].typeIn, '용접 불량');
  assert.equal(p.records[0].action, '재작업');
  assert.equal(p.records[1].problem, '미조임'); // 내용이 비면 유형 칸을 내용으로
  const e = L.enrich(p.records, TYPES);
  assert.equal(e[0].type, '용접불량'); assert.ok(e[0].column); // 띄어쓰기가 달라도 사전 이름에 맞춤
  assert.equal(e[1].type, '미조임');
});
test('필수 열이 없으면 행을 읽지 않고 · 사용자가 열을 고르면 다시 읽음', () => {
  const aoa = [['발생일자', '품번', '불량내용', '거래선'], ['2026-10-05', 'X-1', '균열', '협력사 A']];
  const d = L.detectColumns(aoa);
  assert.deepEqual(d.missing, ['supplier']);
  assert.equal(L.parseSheet(aoa).records.length, 0);
  const p = L.parseSheet(aoa, { overrides: { supplier: 3 } });
  assert.equal(p.records.length, 1);
  assert.equal(p.records[0].supplier, '협력사 A');
});
test('빈 줄·합계 줄 건너뜀 · 날짜 못 읽은 행 수 셈 · 칸 안 줄바꿈 정리', () => {
  const aoa = [['발생일자', '품명', '품번', '문제점', '협력사'], ['2026-10-05', '커버', 'X-1', '도장\n흐름', '협력사 B'], [], ['합계', null, null, null, null], ['미상', '커버', 'X-2', '찍힘', '협력사 B']];
  const p = L.parseSheet(aoa);
  assert.equal(p.records.length, 2);
  assert.equal(p.records[0].problem, '도장 흐름');
  assert.equal(p.badDate, 1);
  assert.equal(p.skipped, 1);
});
test('여러 시트 중 Raw data 시트를 고름', () => {
  const pick = L.pickSheet([{ name: '요약', aoa: [['협력사별 합계'], ['A', 3]] }, { name: 'RAW', aoa: Sample.aoa() }]);
  assert.equal(pick.name, 'RAW');
});
test('같은 행은 한 번만 (발생일자+협력사+품번+문제점, 띄어쓰기 무시)', () => {
  const a = { date: '2026-10-05', supplier: '협력사 A', partNo: 'x-1', problem: '용접 균열' };
  const m1 = L.mergeRecords([], [a], 0);
  const m2 = L.mergeRecords(m1.records, [{ date: '2026-10-05', supplier: '협력사A', partNo: 'X-1', problem: '용접균열' }, { ...a, date: '2026-10-06' }], m1.seq);
  assert.equal(m2.added, 1); assert.equal(m2.dup, 1); assert.equal(m2.records.length, 2);
  assert.equal(m2.records[1].id, 'd2');
});

console.log('캠퍼스(울산 · 인천) RAW DATA — 4차');
test('열 찾기 — 줄 전체에서 이름이 꼭 같은 칸 먼저(왼쪽 「Vendor Code」「문제점 번호」「조치 담당자명」이 가로채지 않음) · 별칭 순서(협력회사명 > 협력회사)', () => {
  const d = L.detectColumns([['Vendor Code', 'Vendor Name', 'Issued Date', 'Problem Description', 'Defect Dev. Desc']]);
  assert.equal(d.map.supplier, 1); assert.equal(d.map.date, 2); assert.equal(d.map.problem, 3);
  const u = L.detectColumns([['문제점 번호', '발생 일자', '협력회사', '협력회사명', '조치 담당자명', '문제점 상세', '조치방안 상세', '조치 방안']]);
  assert.equal(u.map.srcNo, 0); assert.equal(u.map.supplier, 3); assert.equal(u.map.problem, 5); assert.equal(u.map.action, 6);
  assert.deepEqual(u.missing, []);
});
test('울산 · 인천 열 구성이 달라도 같은 항목으로 읽음(제목 줄 · 영문 머리)', () => {
  const u = L.parseSheet(Sample.campusAoa('울산')), i = L.parseSheet(Sample.campusAoa('인천'));
  assert.equal(u.det.headerRow, 2); assert.equal(i.det.headerRow, 0);
  for (const p of [u, i]) { assert.deepEqual(p.det.missing, []); assert.equal(p.badDate, 0); assert.equal(p.skipped, 0); }
  const r0 = Sample.campusRecords('울산')[0];
  assert.deepEqual([u.records[0].date, u.records[0].supplier, u.records[0].partNo, u.records[0].partName, u.records[0].problem, u.records[0].qty, u.records[0].srcNo], [r0.date, r0.supplier, r0.partNo, r0.partName, r0.problem, r0.qty, 'U-26-0001']);
  const k0 = Sample.campusRecords('인천')[0];
  assert.deepEqual([i.records[0].date, i.records[0].supplier, i.records[0].partNo, i.records[0].problem, i.records[0].srcNo], [k0.date, k0.supplier, k0.partNo, k0.problem, 'QN-7001']);
});
test('합치기 — 캠퍼스 칸을 붙이고 중복은 캠퍼스 안에서만 · 다른 캠퍼스의 같은 내용은 남김', () => {
  const U = L.parseSheet(Sample.campusAoa('울산')).records, I = L.parseSheet(Sample.campusAoa('인천')).records;
  let m = L.replaceCampus([], '울산', U, 0);
  assert.equal(m.dup, 1); assert.equal(m.added, U.length - 1);
  m = L.replaceCampus(m.records, '인천', I, m.seq);
  assert.equal(m.dup, 0); assert.equal(m.added, I.length);
  const c = L.campusCounts(m.records);
  assert.deepEqual([c['울산'], c['인천'], c['']], [U.length - 1, I.length, 0]);
  assert.ok(m.records.every(r => r.src === 'raw' && L.CAMPUSES.includes(r.campus)));
  const first = Sample.campusRecords('울산')[0];
  assert.equal(m.records.filter(r => r.problem === first.problem && r.date === first.date && r.partNo === first.partNo).map(r => r.campus).sort().join(','), '울산,인천');
  assert.equal(m.records.length, Sample.records().length + 1);
});
test('수강생이 보낸 울산 RAW 머리 94칸 그대로(문제점 상세 · 문제점 상세2 · 협력회사 · 협력회사명 · 조치 방안 · 조치방안 상세가 함께 있음) — 맞는 칸을 고름', () => {
  const UH = ["저장 상태", "변경 상태", "IF 상태", "문제점 번호", "발생 일자", "월", "생산라인", "품번", "부품 상세", "협력회사", "협력회사명", "양산 품질 관리자명", "조치 담당자명", "조치 담당자", "QM 판정", "QM 판정 상세", "판정 사유", "MES 지표 산입", "NCR 번호", "NCR 반영상태", "문제점 상세", "첨부", "미리보기", "수량", "공수 손실", "문제점 상세2", "공수 손실 상세", "생산오더", "모델", "Serial No.", "조치방안 상세", "품질개선대상", "대책서첨부", "모델 상세", "문제점 구분", "문제점구분명", "문제점 유형", "문제점유형명", "IF 일자", "부적합 유형", "부적합 유형 상세", "IF 상태2", "IF Message", "추정 귀책", "추정 귀책 부서명", "발행자", "발행자명", "플랜트", "대책서승인", "대책서반려사유", "내외자 구분", "공정 1", "공정 1 상세", "공정 2", "공정 2 상세", "귀책그룹", "귀책그룹명칭", "부품 시리얼 번호", "조치 방안", "전송 차수", "NCR 승인상태", "귀책업체", "귀책업체명", "귀책비율", "귀책부서1", "귀책부서1 명", "귀책비율 1", "귀책부서2", "귀책부서2 명", "귀책비율 2", "귀책부서3", "귀책부서3 명", "귀책비율 3", "NCR 지표산입", "APQP 대상", "APQP 대상 사유", "APAP 대상 사유", "불인정 수용여부", "귀책인정", "귀책불인정사유", "귀책재검토결과", "공수인정", "공수불인정사유", "공수재검토결과", "비용인정", "비용불인정사유", "비용재검토결과", "통화", "청구비용", "합의율", "합의비용", "합의일자", "보상방법", "보상방법명칭"];
  const row = UH.map(h => ({ '발생 일자': 46083, '문제점 번호': 'U-1', '협력회사': 'V-101', '협력회사명': '협력사 A', '문제점 상세': '용접부 균열', '문제점 상세2': '다른 글', '부품 상세': '브래킷', '품번': 'TP-1', '수량': 3, '조치 방안': 'A', '조치방안 상세': '재작업', '귀책업체명': '협력사 Z' })[h] ?? null);
  const p = L.parseSheet([['제목'], [], UH, row]);
  const H0 = n => UH.indexOf(n);
  assert.deepEqual([p.det.map.date, p.det.map.supplier, p.det.map.problem, p.det.map.action, p.det.map.partNo, p.det.map.partName, p.det.map.srcNo, p.det.map.qty],
    [H0('발생 일자'), H0('협력회사명'), H0('문제점 상세'), H0('조치방안 상세'), H0('품번'), H0('부품 상세'), H0('문제점 번호'), H0('수량')]);
  assert.deepEqual([p.records[0].date, p.records[0].supplier, p.records[0].problem, p.records[0].action], ['2026-03-02', '협력사 A', '용접부 균열', '재작업']);
});
test('한 캠퍼스를 다시 올리면 그 캠퍼스 RAW 만 바뀜(다른 캠퍼스 · 직접 등록은 그대로) · 고른 불량유형은 이어 받음', () => {
  const U = L.parseSheet(Sample.campusAoa('울산')).records, I = L.parseSheet(Sample.campusAoa('인천')).records;
  let m = L.replaceCampus([], '울산', U, 0); m = L.replaceCampus(m.records, '인천', I, m.seq);
  const manual = { id: 'd999', date: '2026-10-20', supplier: '협력사 A', partNo: 'TP-1', partName: '', problem: '직접 적은 불량', campus: '울산' };
  const db = m.records.concat([manual]);
  const pick = db.find(r => r.campus === '울산' && r.src === 'raw'); pick.typeManual = '기타';
  const U2 = U.slice(0, 10);
  const m2 = L.replaceCampus(db, '울산', U2, m.seq);
  const c = L.campusCounts(m2.records);
  assert.equal(m2.removed, U.length - 1);
  assert.equal(c['인천'], I.length);
  assert.equal(c['울산'], 10 + 1); // 새 RAW 10건 + 직접 등록 1건
  assert.ok(m2.records.some(r => r.id === 'd999'));
  assert.equal(m2.carried, U2.some(r => L.recordKey({ ...r, campus: '울산' }) === L.recordKey(pick)) ? 1 : 0);
  assert.equal(m2.records.find(r => L.recordKey(r) === L.recordKey(pick)).typeManual, '기타');
  // 되돌려 확인: 캠퍼스를 키에서 빼면(옛 키) 인천의 같은 내용 줄이 중복으로 사라져야 함 — 키에 캠퍼스가 들어 있어 남는다
  assert.notEqual(L.recordKey({ ...pick, campus: '울산' }), L.recordKey({ ...pick, campus: '인천' }));
});
test('분석·Excel 의 캠퍼스 거르기 · 불량목록 시트에 캠퍼스 · 관리번호 칸 · 백업으로 캠퍼스 유지', () => {
  const U = L.parseSheet(Sample.campusAoa('울산')).records, I = L.parseSheet(Sample.campusAoa('인천')).records;
  let m = L.replaceCampus([], '울산', U, 0); m = L.replaceCampus(m.records, '인천', I, m.seq);
  const all = L.enrich(m.records, TYPES);
  const a = L.analyze(all, {}, {}), u = L.analyze(all, { campus: '울산' }, {}), i = L.analyze(all, { campus: '인천' }, {});
  assert.equal(u.total + i.total, a.total); assert.equal(i.total, I.length);
  assert.deepEqual([a.campuses['울산'], a.campuses['인천']], [u.total, i.total]);
  const sh = L.reportSheets({ suppliers: L.suppliersOf(all), records: all, plans: {}, rules: {}, today: '2026-10-27', types: TYPES, campus: '인천' });
  assert.equal(sh['불량목록'].length, I.length + 1);
  assert.ok(sh['불량목록'].slice(1).every(r => r[1] === '인천' && /^QN-/.test(r[10])));
  const back = L.parseSheet(L.backupSheet(m.records)).records;
  assert.deepEqual(back.map(L.recordKey), m.records.map(L.recordKey));
  const db = L.restoreDb({ schemaVersion: 4, records: m.records });
  assert.deepEqual(L.campusCounts(db.records), L.campusCounts(m.records));
  assert.ok(db.records.every(r => r.src === 'raw'));
});

console.log('불량유형 분류');
test('기본 사전 = 수강생이 정한 불량유형 10종(8종 + 도장/외관불량 · 녹/발청, 기타는 맨 끝)', () => {
  assert.deepEqual(L.typeNames(TYPES), ['위치/치수', '간섭/정렬', '이종/누락', '용접불량', '누유/누수', '미조임', '성능이상', '도장/외관불량', '녹/발청', '기타']);
  assert.ok(TYPES.every(t => t.actions.length >= 3 && t.keywords.length > 0));
});
test('키워드로 분류 · 맞은 키워드를 근거로 남김', () => {
  const c = L.classify('브래킷 용접부 갈라짐', TYPES);
  assert.equal(c.type, '용접불량');
  assert.deepEqual(c.hits, ['용접', '갈라짐']);
  assert.equal(L.classify('이종품 혼입 (타 품번 섞임)', TYPES).type, '이종/누락');
  assert.equal(L.classify('와셔 누락', TYPES).type, '이종/누락');
  assert.equal(L.classify('홀 위치 치수 공차 초과', TYPES).type, '위치/치수');
  assert.equal(L.classify('조립 시 상대 부품과 간섭', TYPES).type, '간섭/정렬');
  assert.equal(L.classify('피팅부 오일 샘', TYPES).type, '누유/누수');
  assert.equal(L.classify('호스 클램프 미조임', TYPES).type, '미조임');
  assert.equal(L.classify('밸브 작동 불량', TYPES).type, '성능이상');
  assert.equal(L.classify('도장 흐름 발생', TYPES).type, '도장/외관불량');
  assert.equal(L.classify('표면 찍힘', TYPES).type, '도장/외관불량');
  assert.equal(L.classify('보관 중 녹 발생', TYPES).type, '녹/발청');
  assert.equal(L.classify('나사부 부식', TYPES).type, '녹/발청');
  assert.equal(L.classify('이물 부착', TYPES).type, '기타');
});
test('여러 유형이 걸리면 긴 키워드 → 같으면 사전에서 위에 있는 유형', () => {
  assert.equal(L.classify('이종 볼트 혼용', TYPES).type, '이종/누락'); // 혼용·볼트 모두 2글자 → 위의 이종/누락
  assert.equal(L.classify('볼트 체결 토크 부족', TYPES).type, '미조임');
  assert.equal(L.classify('도막 두께 부족', TYPES).type, '도장/외관불량'); // 「도막 두께」(4글자)가 위치/치수의 「두께」를 이김
  assert.equal(L.classify('Crack 발생', TYPES).type, '용접불량'); // 대소문자 무시
  const custom = [{ name: '외관', keywords: ['찍힘'] }, { name: '포장', keywords: ['포장 찍힘'] }];
  assert.equal(L.classify('포장 찍힘 발견', custom).type, '포장'); // 긴 키워드(띄어쓰기 빼고 4글자)가 이김
});
test('맞는 키워드가 없으면 미분류 · 직접 고른 유형이 먼저 · 유형 이름은 구분 기호 무시', () => {
  assert.equal(L.classify('검사성적서 미첨부', TYPES).type, L.UNCLASSIFIED);
  const r = { problem: '검사성적서 미첨부', typeManual: '기타' };
  assert.deepEqual(L.typeOf(r, TYPES), { type: '기타', hits: [], manual: true });
  assert.equal(L.typeOf({ problem: '용접 기공', typeManual: '사전에 없는 유형' }, TYPES).type, '용접불량');
  assert.equal(L.findTypeName('위치·치수', TYPES), '위치/치수');
  assert.equal(L.findTypeName('누유 / 누수', TYPES), '누유/누수');
  assert.equal(L.findTypeName('용접 불량', TYPES), '용접불량');
  assert.equal(L.findTypeName('도장/외관 불량', TYPES), '도장/외관불량');
  assert.equal(L.findTypeName('녹·발청', TYPES), '녹/발청');
  assert.equal(L.findTypeName('모르는 코드', TYPES), '');
  assert.equal(L.typeOf({ problem: '도장 흐름', typeIn: '모르는 코드' }, TYPES).type, '도장/외관불량'); // 맞는 이름이 없으면 키워드로
});
test('설정 화면 글 ↔ 대책 목록 · 키워드 쉼표 나누기', () => {
  const a = L.parseActionLines('원인 조사 | 협력사 | 0 | 2\n\n검사 강화 | | x | 0\n유효성 확인', '부품품질팀');
  assert.deepEqual(a, [{ text: '원인 조사', dept: '협력사', start: 0, weeks: 2 }, { text: '검사 강화', dept: '부품품질팀', start: 0, weeks: 2 }, { text: '유효성 확인', dept: '부품품질팀', start: 0, weeks: 2 }]);
  assert.deepEqual(L.parseActionLines(L.formatActionLines(a)), a);
  assert.deepEqual(L.parseKeywords('균열, 크랙,\n균열 , crack'), ['균열', '크랙', 'crack']);
});

console.log('분석·개선 필요 항목');
const REC = L.mergeRecords([], Sample.records(), 0).records;
const ALL = L.enrich(REC, TYPES);
test('예시 데이터 — 협력사 6곳, 미분류 3건, 파레토 누적 100%', () => {
  const an = L.analyze(ALL, {}, {});
  assert.equal(an.suppliers.length, 6);
  assert.equal(an.unclassified, 3);
  assert.equal(an.byType[an.byType.length - 1].type, L.UNCLASSIFIED); // 미분류는 맨 뒤
  assert.equal(an.byType[an.byType.length - 1].cum, 100);
  assert.equal(an.byType.reduce((s, t) => s + t.count, 0), an.total);
  assert.equal(an.months[0].month, '2026-01');
  assert.equal(an.months[an.months.length - 1].month, '2026-10');
});
test('월별 추이는 빈 달도 0 으로 채움 · 기간 거르기', () => {
  const rows = L.enrich([{ date: '2026-01-10', supplier: 'S', partNo: 'P', problem: '찍힘' }, { date: '2026-04-02', supplier: 'S', partNo: 'P', problem: '찍힘' }], TYPES);
  const an = L.analyze(rows, {}, {});
  assert.deepEqual(an.months.map(m => m.count), [1, 0, 0, 1]);
  assert.equal(L.analyze(rows, { from: '2026-02-01' }, {}).total, 1);
});
test('개선 필요 ① 불량유형 WORST 1·2·3 — 협력사마다 건수 순, 미분류 제외', () => {
  const an = L.analyze(ALL, { supplier: '협력사 C' }, {});
  const w = an.issues.filter(i => i.kind === 'worst').sort((a, b) => a.rank - b.rank);
  assert.deepEqual(w.map(i => i.label + ' ' + i.type), ['WORST 1 용접불량', 'WORST 2 위치/치수', 'WORST 3 간섭/정렬']);
  assert.ok(w[0].count >= w[1].count && w[1].count >= w[2].count);
  // 전체로 보면 협력사마다 따로 순위(협력사 6곳 × 최대 3)
  const all = L.analyze(ALL, {}, {});
  assert.equal(all.worst.length, 17); // F 는 분류된 유형이 2개뿐
  assert.ok(all.worst.every(i => i.type !== L.UNCLASSIFIED));
});
test('개선 필요 ② 동일품번 2건 이상 — 유형이 달라도 셈 · 심어 둔 TP-3301·TP-1102·TP-4405', () => {
  const c = L.analyze(ALL, { supplier: '협력사 C' }, {}).issues.find(i => i.kind === 'part' && i.partNo === 'TP-3301');
  assert.ok(c && c.count >= 6 && c.types.includes('용접불량'));
  assert.ok(L.analyze(ALL, { supplier: '협력사 A' }, {}).issues.some(i => i.kind === 'part' && i.partNo === 'TP-1102'));
  assert.ok(L.analyze(ALL, { supplier: '협력사 D' }, {}).issues.some(i => i.kind === 'part' && i.partNo === 'TP-4405'));
});
test('1건이어도 WORST 에 들어감(협력사 F 3건) · 규칙 경계', () => {
  const f = L.analyze(ALL, { supplier: '협력사 F' }, {}).issues;
  assert.deepEqual(f.map(i => i.label), ['WORST 1', 'WORST 2']);
  const mk = (list) => list.map((x, i) => ({ date: '2026-0' + (1 + i % 3) + '-1' + i, supplier: 'S', partNo: x[1], problem: x[0] }));
  const recs = L.enrich(mk([['찍힘', 'P1'], ['찍힘', 'P2'], ['와셔 누락', 'P3'], ['오일 샘', 'P4'], ['볼트 풀림', 'P3'], ['성적서 문의', 'P5']]), TYPES);
  const an = L.analyze(recs, {}, {});
  // 2건 → 1건끼리는 최근 발생일 순: 이종/누락(03-12) · 미조임(02-14) · 누유/누수(01-13). 사전 순서(이종·누유·미조임)와 일부러 다르게
  assert.deepEqual(an.issues.filter(i => i.kind === 'worst').map(i => i.type), ['도장/외관불량', '이종/누락', '미조임']);
  assert.deepEqual(an.issues.filter(i => i.kind === 'part').map(i => i.partNo), ['P3']); // 같은 품번 2건(유형이 달라도)
  assert.equal(L.analyze(recs, {}, { worstN: 1, partRepeatMin: 3 }).issues.length, 1);
  assert.deepEqual(L.cleanRules({ worstN: 0, partRepeatMin: 1 }), { worstN: 3, partRepeatMin: 2 });
});

console.log('Master Plan 초안·간트');
test('초안 — WORST 1·2·3 유형마다 기본 대책(1주씩 늦춤) + 동일품번마다 재발 방지 한 줄', () => {
  const an = L.analyze(ALL, { supplier: '협력사 C' }, {});
  const rows = L.buildDraft(an, TYPES, { supplier: '협력사 C', baseDate: '2026-10-26', newId });
  const tpl = rows.filter(r => !r.action.startsWith('동일품번'));
  assert.deepEqual([...new Set(tpl.map(r => r.type))], ['용접불량', '위치/치수', '간섭/정렬']);
  const crack = rows.filter(r => r.type === '용접불량' && !r.action.startsWith('동일품번'));
  assert.equal(crack.length, TYPES.find(t => t.name === '용접불량').actions.length);
  assert.equal(crack[0].planStart, '2026-10-26');
  assert.equal(crack[0].planEnd, L.addDays(crack[0].planStart, 13)); // 2주
  assert.equal(rows.find(r => r.type === '위치/치수').planStart, '2026-11-02'); // 둘째 유형은 1주 뒤
  const parts = rows.filter(r => r.action.startsWith('동일품번'));
  assert.equal(parts.length, an.issues.filter(i => i.kind === 'part').length);
  assert.ok(parts[0].action.includes('TP-3301'));
  assert.equal(parts[0].type, '용접불량'); // 품번의 대표 유형 = 가장 많이 나온 유형
  assert.ok(rows.every(r => r.status === 'open' && r.source === 'template' && r.supplier === '협력사 C' && r.note === '')); // 근거 글은 비고에 넣지 않음
  assert.ok(rows.every(r => L.DEFAULT_DEPTS.includes(r.dept)), '담당부서는 협력사 내부 부서 5종');
});
test('담당부서 기본 = 품질 · 생산기술 · 영업 · 설계 · 구매 (기본 대책에 모두 쓰임)', () => {
  assert.deepEqual(L.DEFAULT_DEPTS, ['품질', '생산기술', '영업', '설계', '구매']);
  const used = new Set(TYPES.flatMap(t => t.actions.map(a => a.dept)));
  assert.deepEqual([...used].sort(), [...L.DEFAULT_DEPTS].sort());
});
test('진행상태(4차) = 진행중 · 완료 · 미완료 · 보류 — 마감이 지나면 미완료, 시작 전·날짜 없음은 진행중', () => {
  const r = (o) => Object.assign({ status: 'open', planStart: '2026-11-02', planEnd: '2026-11-15', actStart: '' }, o);
  const T0 = '2026-11-10';
  assert.equal(L.phaseOf(r({ status: 'done' }), T0), 'done');
  assert.equal(L.phaseOf(r({ status: 'hold', planEnd: '2026-11-01' }), T0), 'hold'); // 마감이 지나도 보류는 보류
  assert.equal(L.phaseOf(r({ planEnd: '2026-11-09' }), T0), 'late');
  assert.equal(L.phaseOf(r({}), T0), 'doing');
  assert.equal(L.phaseOf(r({}), '2026-11-02'), 'doing'); // 시작일 당일
  assert.equal(L.phaseOf(r({}), '2026-11-15'), 'doing'); // 마감일 당일
  assert.equal(L.phaseOf(r({}), '2026-11-16'), 'late');  // 마감 다음 날 = 미완료
  assert.equal(L.phaseOf(r({}), '2026-11-01'), 'doing'); // 시작 전도 진행중(「시작 전」은 덧붙임 표시만)
  assert.equal(L.notStarted(r({}), '2026-11-01'), true); assert.equal(L.notStarted(r({}), T0), false);
  assert.equal(L.notStarted(r({ actStart: '2026-10-30' }), '2026-11-01'), false); // 먼저 시작했으면 시작 전 아님
  assert.equal(L.phaseOf(r({ planStart: '', planEnd: '' }), T0), 'doing');
  assert.deepEqual(L.PHASE_ORDER.map(k => L.PHASE[k]), ['진행중', '완료', '미완료', '보류']);
  assert.deepEqual(L.PHASE_IN, ['진행중', '완료', '미완료', '보류']);
  assert.ok(!('progressOf' in L) && !('cleanProgress' in L), '진행률 없음');
  assert.ok(!L.makeRow({ supplier: 'S', action: 'a' }, newId).hasOwnProperty('progress'));
});
test('추진 집계 — 전체 · 진행중 · 완료 · 미완료 · 보류 · 추진 기간 일수', () => {
  const rows = [
    { status: 'done', planStart: '2026-10-26', planEnd: '2026-11-08' },
    { status: 'open', planStart: '2026-11-02', planEnd: '2026-11-15' },
    { status: 'hold', planStart: '2026-11-09', planEnd: '2026-12-20' },
    { status: 'open', planStart: '2026-11-30', planEnd: '2026-12-13' },
    { status: 'open', planStart: '2026-11-16', planEnd: '2026-11-29' }
  ];
  const s = L.planStats(rows, '2026-11-20');
  assert.deepEqual([s.total, s.done, s.doing, s.late, s.hold, s.notStarted, s.overdue, s.days, s.rate], [5, 1, 2, 1, 1, 1, 1, 56, 20]);
  assert.equal(s.doing + s.done + s.late + s.hold, s.total, '네 가지 합 = 전체');
  assert.ok(!('planned' in s));
});
test('간트 주 칸 — 월요일부터 · 위 줄은 달 · 칸 걸침', () => {
  const rows = [{ planStart: '2026-10-28', planEnd: '2026-11-10' }];
  const cols = L.ganttColumns(rows, 'week', { minCols: 1 });
  assert.equal(cols[0].start, '2026-10-26');
  assert.equal(cols[0].full, '2026년 10월 5주차');
  assert.equal(cols.length, 3);
  assert.deepEqual(L.groupSpans(cols), [{ group: '2026년 10월', n: 1 }, { group: '2026년 11월', n: 2 }]);
  assert.deepEqual(L.span('2026-10-28', '2026-11-10', cols), [0, 2]);
  assert.deepEqual(L.span('2026-11-03', '', cols), [1, 1]); // 끝이 비면 시작 칸 하나
  assert.equal(L.span('', '2026-11-03', cols), null);
  assert.equal(L.ganttColumns(rows, 'week', {}).length, 10); // 최소 10칸
});
test('간트 월 칸 — 위 줄은 연도 · 해를 넘김', () => {
  const cols = L.ganttColumns([{ planStart: '2026-11-15', planEnd: '2027-02-03' }], 'month', { minCols: 1 });
  assert.deepEqual(cols.map(c => c.label), ['11월', '12월', '1월', '2월']);
  assert.deepEqual(L.groupSpans(cols).map(g => g.group + g.n), ['2026년2', '2027년2']);
  assert.equal(cols[1].end, '2026-12-31');
});
test('입력 점검 — 날짜 순서 · 완료에는 실적 완료일', () => {
  assert.deepEqual(L.checkPlanRow({ action: 'a', planStart: '2026-11-02', planEnd: '2026-11-01', status: 'open' }), ['계획 완료일이 시작일보다 빠릅니다.']);
  assert.equal(L.checkPlanRow({ action: 'a', status: 'done' }).length, 1);
  assert.equal(L.checkPlanRow({ action: 'a', status: 'done', actStart: '2026-11-02', actEnd: '2026-11-05' }).length, 0);
});
test('계획 확정 — 버전 쌓기 · 실적 입력은 수정으로 보지 않고 계획일을 바꾸면 수정 있음', () => {
  const plan = { rows: [L.makeRow({ supplier: 'S', type: '기타', action: '포장 개선', planStart: '2026-11-02', planEnd: '2026-11-15' }, newId)] };
  assert.equal(L.planStateText(plan), '미확정 초안');
  const s1 = L.confirmPlan(plan, '2026-10-20T10:30:00');
  assert.equal(s1.version, 1);
  plan.rows[0].actStart = '2026-11-03'; plan.rows[0].status = 'done';
  assert.equal(L.planChangedSince(plan), false);
  assert.equal(s1.rows[0].actStart, ''); // 확정본은 복사본
  plan.rows[0].planEnd = '2026-11-22';
  assert.equal(L.planChangedSince(plan), true);
  assert.equal(L.planStateText(plan), '확정 v1 (2026-10-20 10:30) 이후 계획 수정 있음');
  assert.equal(L.confirmPlan(plan, '2026-10-21T09:00:00').version, 2);
});

console.log('AI 프롬프트·답 읽기');
const AL = L.updateAliases({}, REC);
test('별칭 — 같은 이름은 늘 같은 별칭 · 새 이름에만 다음 번호', () => {
  assert.equal(AL.suppliers['협력사 A'], '업체1');
  const again = L.updateAliases(AL, REC.concat([{ supplier: '협력사 Z', partNo: 'NEW-1' }]));
  assert.equal(again.suppliers['협력사 A'], '업체1');
  assert.equal(again.suppliers['협력사 Z'], '업체7');
  assert.equal(again.parts['NEW-1'], '품목' + (Object.keys(AL.parts).length + 1));
});
test('프롬프트 — 협력사·품번이 남지 않음(문제점 글 속 품번 포함) · 사실 숫자 포함', () => {
  const an = L.analyze(ALL, { supplier: '협력사 C' }, {});
  const rows = L.buildDraft(an, TYPES, { supplier: '협력사 C', baseDate: '2026-10-26', newId });
  const info = L.buildAiPromptInfo({ supplier: '협력사 C', an, rows, types: TYPES, depts: L.DEFAULT_DEPTS, baseDate: '2026-10-26', aliases: AL });
  assert.deepEqual(info.leaks, []);
  assert.ok(!/협력사 [A-F]/.test(info.prompt));
  assert.ok(!/TP-\d{4}/.test(info.prompt));
  assert.ok(info.prompt.includes('협력사: 업체3'));
  assert.ok(info.prompt.includes('공정불량 ' + an.total + '건'));
  assert.ok(info.prompt.includes('| 불량유형 | 개선대책 | 담당부서 |'));
  assert.ok(info.prompt.includes('WORST 1 용접불량') && info.prompt.includes('동일품번 2건 이상'));
  assert.ok(info.prompt.includes('담당부서(협력사 내부 부서) 목록: 품질, 생산기술, 영업, 설계, 구매'));
  const short = L.buildAiPromptInfo({ supplier: '협력사 C', an, rows, types: TYPES, depts: L.DEFAULT_DEPTS, baseDate: '2026-10-26', aliases: AL, compact: true });
  assert.ok(short.prompt.length < info.prompt.length && !short.prompt.includes('지금 초안'), '짧게: ' + short.prompt.length + ' < ' + info.prompt.length);
  assert.deepEqual(L.AI_TARGETS.map(t => t.id), ['copilot', 'chatgpt', 'claude', 'gemini']);
  const open = L.buildAiPromptInfo({ supplier: '협력사 C', an, rows, types: TYPES, baseDate: '2026-10-26', aliases: null });
  assert.ok(open.prompt.includes('협력사 C') && open.prompt.includes('TP-3301'));
  assert.ok(L.maskText('TP-3301 재발, TP-33011 은 다른 품번', AL).startsWith(AL.parts['TP-3301'] + ' 재발, TP-33011'));
});
const P3301 = AL.parts['TP-3301'];
test('AI 답 읽기 — 요약·표·확인 필요 · 별칭 되돌림 · 유형 이름 맞춤 · 주 → 날짜', () => {
  const ans = [
    '**요약:** 업체3 은 용접불량이 최근 늘었고 ' + P3301 + ' 에서 반복돼.',
    '',
    '| 불량유형 | 개선대책 | 담당부서 | 시작(기준일부터 몇 주 뒤) | 기간(주) |',
    '|:---|---|---|---|---|',
    '| 용접 불량 | ' + P3301 + ' 용접 예열 온도 기록 시작 | 생산기술팀 | 0 | 2주 |',
    '| 위치·치수 | PT 검사<br>샘플링 2배 | 품질 | 1주 | 4 |',
    '| 새 유형 | 기타 대책 | | | |',
    '| 잘린 줄 |',
    '확인 필요:',
    '- 예열 설비가 있는지',
    '- 품목 3개 공정 순서'
  ].join('\n');
  const p = L.parseAiPlan(ans, { types: TYPES, aliases: AL, supplier: '협력사 C', baseDate: '2026-10-26', newId });
  assert.equal(p.summary, '협력사 C 은 용접불량이 최근 늘었고 TP-3301 에서 반복돼.');
  assert.equal(p.rows.length, 3);
  assert.equal(p.skipped, 1);
  assert.equal(p.rows[0].type, '용접불량');
  assert.equal(p.rows[1].type, '위치/치수'); // 「위치·치수」 → 사전 이름
  assert.equal(p.rows[0].action, 'TP-3301 용접 예열 온도 기록 시작');
  assert.equal(p.rows[0].planEnd, '2026-11-08');
  assert.equal(p.rows[1].action, 'PT 검사 샘플링 2배');
  assert.equal(p.rows[1].planStart, '2026-11-02');
  assert.equal(p.rows[2].type, '새 유형');
  assert.equal(p.rows[2].dept, '품질');
  assert.equal(p.rows[0].dept, '생산기술'); // 「생산기술팀」 → 목록의 이름
  assert.ok(p.rows.every(r => r.source === 'ai' && r.supplier === '협력사 C'));
  assert.deepEqual(p.checks, ['예열 설비가 있는지', '품목 3개 공정 순서']); // 「품목 3개」는 별칭이 아님
});
test('엑셀에서 복사한 탭 구분 줄도 읽음', () => {
  const p = L.parseAiPlan('기타\t도막 측정 추가\t품질\t2\t3', { types: TYPES, aliases: {}, supplier: 'S', baseDate: '2026-10-26', newId });
  assert.equal(p.rows.length, 1);
  assert.equal(p.rows[0].planStart, '2026-11-09');
  assert.equal(p.rows[0].planEnd, '2026-11-29');
});
test('AI 연결 — 키 없이 보내면 인증 머리글 없음 · https 페이지에서 http 주소 막음', () => {
  const r = E.buildChatRequest({ baseUrl: 'https://llm.example.com/v1/chat/completions/', model: 'm' }, E.promptMessages('x'));
  assert.equal(r.url, 'https://llm.example.com/v1/chat/completions');
  assert.equal(r.init.headers.Authorization, undefined);
  assert.ok(E.validateConfig({ baseUrl: 'http://10.0.0.5:8000/v1', model: 'm' }, 'https:').errors.length > 0);
  assert.equal(E.toStored({ baseUrl: 'https://a/v1', model: 'm', apiKey: 'sk-x', remember: false }).apiKey, '');
});

console.log('보고서·저장');
const SHEETS = ['불량목록', 'MASTER PLAN'];
function reportFixture() {
  const an = L.analyze(ALL, { supplier: '협력사 C' }, {});
  const rows = L.buildDraft(an, TYPES, { supplier: '협력사 C', baseDate: '2026-10-26', newId });
  rows[0].actStart = '2026-10-26'; rows[0].actEnd = '2026-10-30'; rows[0].status = 'done';
  rows[1].status = 'hold';
  const plan = { baseDate: '2026-10-26', from: '', to: '', rows, confirmed: [] };
  L.confirmPlan(plan, '2026-10-20T11:00:00');
  const plans = { '협력사 C': plan };
  return { an, rows, plan, plans, book: L.reportBook({ suppliers: ['협력사 C'], records: ALL, plans, rules: {}, today: '2026-11-05', types: TYPES }) };
}
const H = 4; // MASTER PLAN 시트의 머리 줄(0부터)
test('Excel 추출은 「불량목록」·「MASTER PLAN」 두 시트 · 개선근거 칸 없음 · 숨긴 ID·기준 칸', () => {
  const { an, rows, book } = reportFixture();
  const sh = book.sheets;
  assert.deepEqual(Object.keys(sh), SHEETS);
  assert.deepEqual(sh['불량목록'][0], ['No', '캠퍼스', '발생일자', '협력사', '품번', '품명', '불량유형', '내용', '수량', '조치', '관리번호']);
  assert.equal(sh['불량목록'].length, an.total + 1);
  const mp = sh['MASTER PLAN'];
  assert.deepEqual(mp[H], ['ID', 'No', '협력사', '불량유형', '개선대책', '담당부서', '담당자', '개선 시작(계획)', '개선 마감(계획)', '실적 시작', '실적 완료', '진행상태', '비고', '기준(지우지 마세요)']);
  assert.ok(!mp[H].some(h => /진행률/.test(h)), '진행률 칸 없음');
  assert.ok(!mp[H].some(h => /근거/.test(h)));
  assert.equal(mp.length, H + 1 + rows.length);
  assert.equal(mp[H + 1][0], rows[0].id);
  assert.equal(mp[H + 1][11], '완료');
  assert.equal(mp[H + 2][11], '보류');
  assert.ok(mp.slice(H + 1).some(r => r[11] === '진행중'));
  assert.ok(mp.slice(H + 1).every(r => ['진행중', '완료', '미완료', '보류'].includes(r[11])), '진행상태는 네 값만');
  assert.deepEqual(JSON.parse(mp[H + 1][13]).status, 'done');
  assert.ok(!('progress' in JSON.parse(mp[H + 1][13])));
  assert.ok(mp[1][0].includes('진행중') && mp[1][0].includes('미완료') && !mp[1][0].includes('예정'));
  const m = book.meta['MASTER PLAN'];
  assert.deepEqual(m.hidden, [0, 13]);
  assert.ok(m.cells.some(c => c.r === H && c.c === 4 && c.style === 'edit'));
  assert.equal(m.lists[0].items.join(','), '진행중,완료,미완료,보류');
  assert.equal(m.lists[0].sqref, 'L' + (H + 2) + ':L' + (H + 500));
  assert.equal(m.cf[0].sqref, m.lists[0].sqref);
  assert.deepEqual(m.cf[0].rules.map(r => r.equal + '=' + r.style), ['완료=done', '진행중=doing', '미완료=late', '보류=hold']);
});
test('Master Plan 이 없는 협력사도 불량목록은 나옴', () => {
  const sh = L.reportSheets({ suppliers: ['협력사 A', '협력사 B'], records: ALL, plans: {}, rules: {}, today: '2026-10-20', types: TYPES });
  assert.equal(sh['불량목록'].length, ALL.filter(r => r.supplier === '협력사 A' || r.supplier === '협력사 B').length + 1);
  assert.equal(sh['MASTER PLAN'].length, H + 1);
});
test('Excel 쓰기 — 머리글 굵게·상태 색·날짜는 엑셀 날짜·틀 고정·숨긴 열·고르기 목록 · 다시 읽으면 값 그대로', () => {
  const { book } = reportFixture();
  const u8 = XO.build(XLSX, book.sheets, book.meta);
  const wb = XLSX.read(u8, { type: 'array', cellStyles: true });
  assert.deepEqual(wb.SheetNames, SHEETS);
  const list = wb.Sheets['불량목록'];
  assert.equal(list.C2.t, 'n'); assert.equal(list.C2.z, 'yyyy-mm-dd');
  assert.equal(list.C2.w, book.sheets['불량목록'][1][2]);
  const z = XLSX.CFB.read(u8, { type: 'array' });
  const xml = p => Buffer.from(z.FileIndex[z.FullPaths.indexOf('Root Entry/' + p)].content).toString('utf8');
  const styles = xml('xl/styles.xml');
  assert.ok(/<font><b\/>/.test(styles), '굵은 글꼴');
  // 조건부 서식(4차) — styles.xml 의 dxf 4개(완료 초록 · 진행중 주황 · 미완료 빨강 · 보류 회색) + 시트의 cfRule 4개
  const dx = /<dxfs count="4">([\s\S]*?)<\/dxfs>/.exec(styles);
  assert.ok(dx, 'dxfs 4개');
  assert.deepEqual([...dx[1].matchAll(/<bgColor rgb="(\w+)"/g)].map(x => x[1]), ['FF1F9D55', 'FFE8830C', 'FFC0392B', 'FF7F8C8D']);
  const s1 = xml('xl/worksheets/sheet1.xml');
  assert.ok(/<c r="A1" s="(\d+)"/.test(s1), '불량목록 머리글에 서식');
  const s2 = xml('xl/worksheets/sheet2.xml');
  assert.ok(s2.includes('state="frozen"') && s2.includes('<autoFilter'), '틀 고정·거르기');
  assert.ok(/<col min="1" max="1"[^>]*hidden="(true|1)"/.test(s2) && /<col min="14" max="14"[^>]*hidden="(true|1)"/.test(s2), 'ID·기준 칸 숨김');
  assert.ok(/<dataValidations count="2">.*진행중,완료,미완료,보류/.test(s2), '진행상태 고르기 목록');
  const cf = /<conditionalFormatting sqref="L6:L504">([\s\S]*?)<\/conditionalFormatting>/.exec(s2);
  assert.ok(cf, '진행상태 칸 조건부 서식');
  assert.deepEqual([...cf[1].matchAll(/<cfRule type="cellIs" dxfId="(\d)" priority="\d" operator="equal"><formula>"(.+?)"<\/formula>/g)].map(x => x[2] + x[1]), ['완료0', '진행중1', '미완료2', '보류3']);
  assert.ok(s2.indexOf('<conditionalFormatting') < s2.indexOf('<dataValidations') && s2.indexOf('</sheetData>') < s2.indexOf('<conditionalFormatting'), 'OOXML 순서: sheetData → conditionalFormatting → dataValidations');
  assert.ok(!new RegExp('<c r="L' + (H + 2) + '" s="\\d+"').test(s2), '진행상태 칸에 박은 색 없음(조건부 서식만)');
  const back = XLSX.utils.sheet_to_json(wb.Sheets['MASTER PLAN'], { header: 1, raw: false });
  assert.equal(back[H + 1][11], '완료');
  assert.equal(back[H + 1][7], book.sheets['MASTER PLAN'][H + 1][7]); // 계획 시작일 글자 그대로
});

console.log('진행사항 업로드(협력사 회신)');
function roundTrip(book, edit) {
  const u8 = XO.build(XLSX, book.sheets, book.meta);
  const wb = XLSX.read(u8, { type: 'array' });
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets['MASTER PLAN'], { header: 1, raw: true, defval: null });
  if (edit) edit(aoa);
  return L.parseProgressSheet(aoa, TYPES);
}
test('회신 읽기 — 바뀐 칸만 · 같은 값은 그대로 · 새 줄 · 모르는 ID · 잘못된 값 · 같은 ID 두 번', () => {
  const { rows, plans, book } = reportFixture();
  const parsed = roundTrip(book, aoa => {
    aoa[H + 3][6] = '김품질 대리';             // 담당자
    aoa[H + 3][9] = '2026-11-02'; aoa[H + 3][11] = '진행중';
    aoa[H + 4][11] = '완료'; aoa[H + 4][10] = 46335; aoa[H + 4][9] = '2026.11.01'; // 엑셀 날짜 일련번호(2026-11-09)
    aoa[H + 5][7] = '날짜 아님';
    aoa.push([null, null, '협력사 C', '용접 불량', '용접 지그 교체', '생산기술', '박생기', '2026-11-16', '2026-11-29', null, null, '진행중', '협력사 추가']);
    aoa.push(['pzzz-none', null, '협력사 C', '기타', '지운 활동', '품질']);
    aoa.push([rows[6].id, null, '협력사 C', rows[6].type, rows[6].action]);
  });
  assert.equal(parsed.ok, true); assert.equal(parsed.hasId, true); assert.equal(parsed.hasBase, true);
  const d = L.diffProgress(plans, parsed, { types: TYPES, knownSuppliers: ['협력사 C'] });
  const at = l => d.items.find(i => i.line === l);
  const ch = at(H + 4);
  assert.equal(ch.kind, 'changed');
  assert.deepEqual(ch.changes.map(c => c.k), ['owner', 'actStart']); // 진행중 = 저장값 open 그대로라 상태는 바뀜 없음
  assert.equal(at(H + 5).kind, 'changed');
  assert.deepEqual(at(H + 5).changes.map(c => c.k + '=' + c.to), ['actStart=2026-11-01', 'actEnd=2026-11-09', 'status=done']);
  assert.equal(at(H + 6).kind, 'invalid'); assert.ok(at(H + 6).reasons[0].includes('날짜'));
  const nw = d.items.find(i => i.kind === 'new');
  assert.ok(nw && nw.vals.type === '용접불량' && nw.vals.planStart === '2026-11-16');
  assert.equal(d.items.find(i => i.id === 'pzzz-none').kind, 'unknown');
  assert.equal(d.items.filter(i => i.id === rows[6].id).map(i => i.kind).join(','), 'same,invalid'); // 두 번째는 같은 ID
  assert.equal(d.counts.same, rows.length - 3 - 1 + 1);
});
test('회신 반영 — 바뀐 칸만 바꾸고 새 줄 더함 · 같은 파일을 다시 올리면 바뀐 것 없음(멱등)', () => {
  const { rows, plans, book } = reportFixture();
  const edit = aoa => {
    aoa[H + 3][6] = '김품질 대리'; aoa[H + 3][9] = '2026-11-03';
    aoa.push([null, null, '협력사 C', '용접불량', '용접 지그 교체', '생산기술', '', '2026-11-16', '2026-11-29']);
  };
  const parsed = roundTrip(book, edit);
  const d = L.diffProgress(plans, parsed, { types: TYPES, knownSuppliers: ['협력사 C'] });
  const n0 = plans['협력사 C'].rows.length;
  const res = L.applyProgress(plans, d, { at: '2026-11-05T10:00:00', newId });
  assert.deepEqual([res.changed, res.added, res.fields], [1, 1, 2]);
  assert.equal(rows[2].owner, '김품질 대리'); assert.equal(rows[2].actStart, '2026-11-03'); assert.equal(rows[2].replyAt, '2026-11-05T10:00:00');
  assert.equal(plans['협력사 C'].rows.length, n0 + 1);
  assert.equal(plans['협력사 C'].rows[n0].source, 'reply');
  const again = L.diffProgress(plans, roundTrip(book, edit), { types: TYPES, knownSuppliers: ['협력사 C'] });
  assert.equal(again.counts.changed + again.counts.new, 0, JSON.stringify(again.counts));
  assert.equal(L.applyProgress(plans, again, { at: 'x', newId }).changed, 0);
  assert.equal(plans['협력사 C'].rows.length, n0 + 1);
});
test('회신 충돌 — 내보낸 뒤 이 도구에서도 고친 칸: 협력사가 안 고쳤으면 도구 값 유지, 둘 다 고쳤으면 충돌(고르면 파일 값)', () => {
  const { rows, plans, book } = reportFixture();
  rows[2].owner = '도구에서 바꿈'; rows[3].note = '도구 메모';
  const parsed = roundTrip(book, aoa => { aoa[H + 4][12] = '협력사 메모'; });
  const d = L.diffProgress(plans, parsed, { types: TYPES });
  assert.equal(d.items.find(i => i.line === H + 4).kind, 'same');        // owner: 파일 = 기준 → 도구 값 유지
  const c = d.items.find(i => i.line === H + 5);
  assert.equal(c.kind, 'changed'); assert.equal(c.changes.length, 0); assert.equal(c.conflicts[0].k, 'note');
  L.applyProgress(plans, d, { at: 't1', newId });
  assert.equal(rows[3].note, '도구 메모');                                // 고르지 않으면 지금 값
  const d2 = L.diffProgress(plans, parsed, { types: TYPES });
  L.applyProgress(plans, d2, { at: 't2', newId, take: { [H + 5]: { note: true } } });
  assert.equal(rows[3].note, '협력사 메모');
  assert.equal(rows[2].owner, '도구에서 바꿈');
});
test('회신 — 완료인데 실적 완료일이 없으면 반영하지 않고 이유를 보임 · 머리 줄 없는 파일은 안내', () => {
  const { plans, book } = reportFixture();
  const d = L.diffProgress(plans, roundTrip(book, aoa => { aoa[H + 3][11] = '완료'; }), { types: TYPES });
  const it = d.items.find(i => i.line === H + 4);
  assert.equal(it.kind, 'invalid'); assert.ok(it.reasons.some(r => r.includes('실적 완료일')));
  assert.equal(L.parseProgressSheet([['아무 표'], ['a', 'b']], TYPES).ok, false);
  assert.equal(L.fingerprint(new Uint8Array([1, 2, 3])), L.fingerprint(new Uint8Array([1, 2, 3])));
});
test('3차에 내보낸 회신 파일(진행률 칸 · 진행 중 · 지연 · 예정)도 읽음 — 진행률은 무시, 세 값은 모두 진행중(저장값 그대로)', () => {
  const { rows, plans } = reportFixture();
  const head = ['ID', 'No', '협력사', '불량유형', '개선대책', '담당부서', '담당자', '개선 시작(계획)', '개선 마감(계획)', '실적 시작', '실적 완료', '진행률(%)', '진행상태', '비고', '기준(지우지 마세요)'];
  const line = (r, st, pg) => [r.id, 1, '협력사 C', r.type, r.action, r.dept, r.owner, r.planStart, r.planEnd, r.actStart, r.actEnd, pg, st, r.note, JSON.stringify({ type: r.type, action: r.action, dept: r.dept, owner: r.owner, planStart: r.planStart, planEnd: r.planEnd, actStart: r.actStart, actEnd: r.actEnd, progress: '', status: r.status, note: r.note })];
  const aoa = [['제목'], [], [], [], head, line(rows[2], '진행 중', 30), line(rows[3], '지연', ''), line(rows[4], '예정', 0)];
  const p = L.parseProgressSheet(aoa, TYPES);
  assert.equal(p.ok, true); assert.ok(!('progress' in p.map));
  const d = L.diffProgress(plans, p, { types: TYPES });
  assert.deepEqual(d.items.map(i => i.kind), ['same', 'same', 'same']);
});
test('저장 형식 — 깨진 값은 버리고 기본값 · 사용자 사전·계획 유지', () => {
  assert.equal(L.restoreDb(null).types.length, 10);
  const db = L.restoreDb({
    records: [{ date: '2026.10.05', supplier: 'S', partNo: 'P', problem: '찍힘', typeManual: '기타', action: '재작업', typeIn: '외관' }, null, 3],
    types: [{ name: '내 유형', keywords: ['가', '가', ''], actions: [{ text: '대책', weeks: 0 }] }, { name: '미분류' }],
    schemaVersion: 3, rules: { worstN: 4 }, mask: false, updates: [{ at: '2026-11-05T10:00', file: 'a.xlsx' }, null],
    plans: { S: { baseDate: '2026-10-26', rows: [{ id: 'x', action: 'a', status: '이상한값', planStart: '2026/11/2' }], confirmed: [{ at: '2026-10-20T10:00', rows: [] }] } }
  });
  assert.equal(db.records.length, 1);
  assert.equal(db.records[0].date, '2026-10-05');
  assert.equal(db.records[0].id, 'd1');
  assert.equal(db.records[0].action, '재작업'); assert.equal(db.records[0].typeIn, '외관');
  assert.deepEqual(db.types.map(t => t.name), ['내 유형']);
  assert.deepEqual(db.types[0].keywords, ['가']);
  assert.equal(db.types[0].actions[0].weeks, 1);
  assert.equal(db.rules.worstN, 4);
  assert.equal(db.updates.length, 1);
  assert.equal(db.mask, false);
  assert.equal(db.plans.S.rows[0].status, 'open');
  assert.equal(db.plans.S.rows[0].planStart, '2026-11-02');
  assert.equal(db.plans.S.confirmed[0].version, 1);
});

test('2차(v2) 저장 데이터 → 3차: 8종에 도장/외관불량·녹/발청을 기타 앞에 · 옛 담당부서는 협력사 내부 부서로(확정본도 같이)', () => {
  const v2 = L.restoreDb({ schemaVersion: 2, depts: ['부품품질팀', '협력사', '구매팀', '생산기술팀', '설계팀'], rules: { repeatMin: 5, partRepeatMin: 3 },
    types: ['위치/치수', '간섭/정렬', '이종/누락', '용접불량', '누유/누수', '미조임', '성능이상', '기타'].map(name => ({ name, dept: '협력사', keywords: name === '기타' ? ['도장', '녹 발생', '이물'] : ['x'], actions: [{ text: '용접 조건(전류·전압·속도·예열) 표준 재설정, 작업표준서 개정', dept: '협력사' }, { text: '유효성 확인(대책 후 재발 여부 추적)', dept: '부품품질팀' }] })),
    plans: { S: { rows: [{ id: 'r1', type: '기타', action: '실링면 가공 상태·오링·피팅 체결 확인', dept: '협력사' }, { id: 'r2', type: '기타', action: 'x', dept: '설계팀' }],
      confirmed: [{ at: '2026-10-20T10:00', rows: [{ id: 'r1', type: '기타', action: '실링면 가공 상태·오링·피팅 체결 확인', dept: '협력사' }, { id: 'r2', type: '기타', action: 'x', dept: '설계팀' }] }] } } });
  assert.deepEqual(L.typeNames(v2.types), L.typeNames(TYPES));
  assert.deepEqual(v2.types[9].keywords, ['이물']);
  assert.deepEqual(v2.depts, L.DEFAULT_DEPTS);
  assert.deepEqual(v2.types[0].actions.map(a => a.dept), ['생산기술', '품질']);
  assert.deepEqual(v2.plans.S.rows.map(r => r.dept), ['생산기술', '설계']);
  assert.equal(L.planChangedSince(v2.plans.S), false);
  assert.deepEqual(v2.rules, L.DEFAULT_RULES);
  // 사용자가 부서를 바꿔 둔 경우는 그대로
  const mine = L.restoreDb({ schemaVersion: 2, depts: ['QC', '생산'], plans: { S: { rows: [{ id: 'r', action: 'a', dept: 'QC' }] } } });
  assert.deepEqual(mine.depts, ['QC', '생산']); assert.equal(mine.plans.S.rows[0].dept, 'QC');
});
test('1단계(v1) 기본 사전으로 저장된 데이터는 10종으로 옮김 · 직접 고른 유형·계획 행도 따라감', () => {
  const V1 = ['이종 불량', '용접 균열', '용접 불량', '가공 치수불량', '도장 불량', '외관 불량', '조립·체결 불량', '누유·누수', '녹·부식'].map(name => ({ name, keywords: ['x'] }));
  const db = L.restoreDb({ schemaVersion: 1, types: V1, records: [{ date: '2026-10-05', supplier: 'S', problem: 'a', typeManual: '가공 치수불량' }],
    plans: { S: { rows: [{ id: 'r1', type: '용접 균열', action: 'a' }], confirmed: [{ at: '2026-10-20T10:00', rows: [{ id: 'r1', type: '조립·체결 불량', action: 'a' }] }] } } });
  assert.deepEqual(L.typeNames(db.types), L.typeNames(TYPES));
  assert.equal(db.records[0].typeManual, '위치/치수');
  assert.equal(db.plans.S.rows[0].type, '용접불량');
  assert.equal(db.plans.S.confirmed[0].rows[0].type, '미조임');
  // 사용자가 이름을 고친 사전은 그대로 둠
  const mine = L.restoreDb({ schemaVersion: 1, types: V1.slice(0, 8).concat([{ name: '내 유형', keywords: ['y'] }]) });
  assert.equal(mine.types[8].name, '내 유형');
});

console.log('직접 입력·엑셀 붙여넣기');
test('붙여넣기 — 머리행 없으면 「발생일자·협력사·품번·품명·불량유형·내용·수량·조치」 순서', () => {
  const t = '2026-10-05\t협력사 A\tTP-1\t커버\t용접 불량\t비드 불량\t3\t재작업\r\n2026.10.06\t협력사 B\t\t\t\t도장 흐름\t\t\n\n';
  const p = L.parsePaste(t);
  assert.equal(p.header, false); assert.equal(p.rows, 2);
  assert.equal(p.records.length, 2);
  assert.deepEqual(p.records[0], { date: '2026-10-05', partName: '커버', partNo: 'TP-1', problem: '비드 불량', supplier: '협력사 A', qty: 3, typeIn: '용접 불량', action: '재작업' });
  assert.equal(p.records[1].date, '2026-10-06');
  assert.equal(L.enrich(p.records, TYPES)[1].type, '도장/외관불량');
});
test('붙여넣기 — 머리행이 있으면 열 이름으로(순서 상관없음) · 칸 안 줄바꿈("…") · 내용 빈 줄은 건너뜀', () => {
  const t = '조치\t불량유형\t업체\t발생일\t내용\n"1차 선별\n2차 반품"\t위치·치수\t협력사 D\t2026-09-01\t"홀 ""A"" 위치 틀림"\n\t\t협력사 D\t2026-09-02\t';
  const p = L.parsePaste(t);
  assert.equal(p.header, true);
  assert.equal(p.records.length, 1); assert.equal(p.skipped, 1);
  assert.equal(p.records[0].action, '1차 선별 2차 반품');
  assert.equal(p.records[0].problem, '홀 "A" 위치 틀림');
  assert.equal(L.enrich(p.records, TYPES)[0].type, '위치/치수');
});
test('붙여넣기 — 데이터 줄을 머리행으로 오해하지 않음(날짜가 든 줄)', () => {
  const p = L.parsePaste('2026-10-05\t협력사 내용\t\t\t\t유형 확인 필요');
  assert.equal(p.header, false);
  assert.equal(p.records[0].supplier, '협력사 내용');
});
test('직접 입력 점검 — 필수(발생일자·협력사·내용) · 수량 숫자 · 유형은 목록에서', () => {
  const ok = { date: '2026-10-05', supplier: '협력사 A', problem: '와셔 누락', type: '이종/누락', qty: '2', action: '선별' };
  assert.deepEqual(L.checkRecordInput(ok, TYPES), []);
  assert.equal(L.checkRecordInput({ ...ok, date: '' }, TYPES).length, 1);
  assert.equal(L.checkRecordInput({ ...ok, supplier: ' ', problem: '' }, TYPES).length, 2);
  assert.equal(L.checkRecordInput({ ...ok, qty: '두 개' }, TYPES).length, 1);
  assert.equal(L.checkRecordInput({ ...ok, type: '없는 유형' }, TYPES).length, 1);
  assert.equal(L.checkRecordInput({ ...ok, type: '' }, TYPES).length, 0); // 비우면 내용 키워드로 자동
  const r = L.recordFromInput(ok);
  assert.deepEqual(r, { date: '2026-10-05', partName: '', partNo: '', problem: '와셔 누락', supplier: '협력사 A', qty: 2, typeManual: '이종/누락', action: '선별' });
  assert.equal(L.enrich([r], TYPES)[0].manual, true);
});
test('엑셀 백업 → 다시 읽으면 같은 레코드(붙여넣기·파일 복원 모두)', () => {
  const recs = REC.slice(0, 20).concat([{ id: 'x', date: '2026-10-05', supplier: '협력사 Z', partNo: '', partName: '', problem: '밸브 작동 불량', typeManual: '성능이상', action: '교체', qty: 1 }]);
  const aoa = L.backupSheet(recs);
  assert.deepEqual(aoa[0], ['발생일자', '협력사', '품번', '품명', '불량유형', '내용', '수량', '조치', '관리번호', '캠퍼스']);
  const back = L.parseSheet(aoa).records;
  assert.equal(back.length, recs.length);
  assert.deepEqual(back.map(L.recordKey), recs.map(L.recordKey));
  assert.equal(L.enrich(back, TYPES)[20].type, '성능이상');
  assert.equal(back[20].action, '교체');
  const pasted = L.parsePaste(aoa.map(r => r.join('\t')).join('\n'));
  assert.equal(pasted.header, true); assert.equal(pasted.records.length, recs.length);
});

console.log('아웃룩 메일 초안(.eml) — 4차');
test('메일 설정 — 기본 수신·참조 3칸 · 협력사별 수신이 있으면 그것 · 빈 칸은 버림 · EMAIL 모양 검사', () => {
  const m = L.cleanMail({ to: [{ name: '김품질', email: 'q1@example.com' }, { name: '', email: '' }], cc: [{ name: '박팀장', email: 'lead@example.com' }],
    per: { '협력사 C': { to: [{ name: 'C 담당', email: 'c@example.co.kr' }], cc: [] }, '협력사 D': { to: [], cc: [] } } });
  assert.equal(m.to.length, 3); assert.equal(m.cc.length, 3);
  assert.deepEqual(Object.keys(m.per), ['협력사 C'], '빈 협력사 칸은 저장하지 않음');
  const c = L.mailFor(m, '협력사 C'), a = L.mailFor(m, '협력사 A');
  assert.deepEqual(c.to.map(p => p.email), ['c@example.co.kr']); assert.deepEqual(c.cc.map(p => p.email), ['lead@example.com']); assert.equal(c.override, true);
  assert.deepEqual(a.to.map(p => p.email), ['q1@example.com']); assert.equal(a.override, false);
  assert.deepEqual(L.checkMail(m, '협력사 A'), []);
  assert.ok(L.checkMail(L.defaultMail(), '협력사 A')[0].includes('받는 사람'));
  assert.ok(L.checkMail(L.cleanMail({ to: [{ email: '이상한 주소' }] }), 'X').some(e => e.includes('EMAIL 모양')));
  assert.deepEqual(L.restoreDb({ mail: m }).mail, m); assert.deepEqual(L.restoreDb({}).mail, L.defaultMail());
});
test('자리표시 — {협력사} · {마감} · {과제수} 등은 값으로, 모르는 {…} 는 그대로', () => {
  const { rows } = reportFixture();
  const v = L.mailVars({ supplier: '협력사 C', due: '2026-11-12', today: '2026-11-05', rows, fileName: 'a.xlsx', fromName: '이품질' });
  assert.equal(v['과제수'], rows.length);
  assert.equal(L.fillTemplate('{협력사} · {마감} · {과제수}건 · {없는값}', v), '협력사 C · 2026-11-12 · ' + rows.length + '건 · {없는값}');
  const s = L.fillTemplate(L.DEFAULT_MAIL.subject, v), b = L.fillTemplate(L.DEFAULT_MAIL.body, v);
  assert.ok(!/\{[^}]+\}/.test(s + b), '기본 제목·본문의 자리표시는 모두 채워짐');
});
test('RFC 2047 — 한글 제목은 75자 이하 낱말로 나누고 글자를 반으로 자르지 않음 · base64 는 표준과 같음', () => {
  const w = L.encodeWord('[공정불량 개선] 협력사 C 개선 MASTER PLAN 회신 요청 (2026-11-12까지) 😀');
  const words = w.split('\r\n ');
  assert.ok(words.length > 1 && words.every(x => x.length <= 75 && /^=\?UTF-8\?B\?[A-Za-z0-9+/=]+\?=$/.test(x)));
  const back = words.map(x => Buffer.from(x.slice(10, -2), 'base64').toString('utf8'));
  assert.ok(back.every(x => !x.includes('�')), '낱말마다 온전한 UTF-8');
  assert.equal(back.join(''), '[공정불량 개선] 협력사 C 개선 MASTER PLAN 회신 요청 (2026-11-12까지) 😀');
  assert.equal(L.encodeWord('Plain ASCII'), 'Plain ASCII');
  for (const n of [0, 1, 2, 3, 100]) { const u = new Uint8Array(n).map((_, i) => (i * 37) & 255); assert.equal(L.base64(u), Buffer.from(u).toString('base64')); }
});
test('.eml — multipart/mixed · X-Unsent: 1 · To/Cc · 첨부 base64 · 줄은 CRLF 76자 · mailto 는 첨부 없이', () => {
  const att = { name: 'MasterPlan_협력사C_2026-11-05.xlsx', type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', bytes: new Uint8Array([80, 75, 3, 4, 1, 2, 3]) };
  const eml = L.buildEml({ to: [{ name: '김품질', email: 'q1@example.com' }], cc: [{ name: 'Lead', email: 'lead@example.com' }], subject: '제목', body: '첫 줄\n둘째 줄', attachments: [att], boundary: 'BND', date: new Date(2026, 10, 5, 9, 0, 0) });
  assert.ok(/^To: =\?UTF-8\?B\?6rmA7ZKI7KeI\?= <q1@example\.com>\r$/m.test(eml));
  assert.ok(/^Cc: "Lead" <lead@example\.com>\r$/m.test(eml));
  assert.ok(eml.includes('\r\nX-Unsent: 1\r\n') && eml.includes('Content-Type: multipart/mixed; boundary="BND"'));
  assert.ok(!/^From:/m.test(eml), '보내는 사람을 비우면 From 없음(아웃룩이 기본 계정으로 채움)');
  assert.ok(eml.includes("filename*=UTF-8''MasterPlan_%ED%98%91%EB%A0%A5%EC%82%ACC_2026-11-05.xlsx"));
  assert.ok(eml.includes('\r\n' + Buffer.from(att.bytes).toString('base64') + '\r\n--BND--'));
  assert.ok(/^[\x00-\x7f]*$/.test(eml), 'ASCII 만');
  assert.ok(eml.split('\r\n').every(l => l.length <= 998) && !eml.replace(/\r\n/g, '').includes('\n'));
  const mt = L.mailtoUrl({ to: [{ email: 'a@x.com' }, { email: 'b@x.com' }], cc: [{ email: 'c@x.com' }], subject: '제목 A', body: '1\n2' });
  assert.equal(mt, 'mailto:a%40x.com,b%40x.com?cc=c%40x.com&subject=%EC%A0%9C%EB%AA%A9%20A&body=1%0D%0A2');
});
test('.eml 을 파이썬 email 패키지로 다시 읽음 — 머리글 풀림 · 첨부가 엑셀 원본과 바이트까지 같음 · openpyxl 로 조건부 서식 4개 (scripts/check-mail.py)', () => {
  const fs = require('node:fs'), os = require('node:os'), path = require('node:path'), cp = require('node:child_process');
  const py = cp.spawnSync('python3', ['-c', 'import email'], { encoding: 'utf8' });
  if (py.status !== 0) { console.log('       (python3 없음 — 건너뜀)'); return; }
  const { rows, book } = reportFixture();
  const xlsx = XO.build(XLSX, book.sheets, book.meta);
  const name = 'MasterPlan_협력사C_2026-11-05.xlsx';
  const mail = L.cleanMail({ to: [{ name: '김품질 책임', email: 'q1@example.com' }, { name: 'Kim, Q', email: 'q2@example.com' }], cc: [{ name: '박팀장', email: 'lead@example.co.kr' }] });
  const v = L.mailVars({ supplier: '협력사 C', due: '2026-11-12', today: '2026-11-05', rows, fileName: name, fromName: '이품질' });
  const subject = L.fillTemplate(mail.subject, v), body = L.fillTemplate(mail.body, v), r = L.mailFor(mail, '협력사 C');
  const eml = L.buildEml({ to: r.to, cc: r.cc, subject, body, attachments: [{ name, type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', bytes: xlsx }] });
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'd1007-'));
  fs.writeFileSync(path.join(dir, 'm.eml'), eml, 'latin1'); fs.writeFileSync(path.join(dir, 'a.xlsx'), xlsx);
  fs.writeFileSync(path.join(dir, 'w.json'), JSON.stringify({ to: r.to, cc: r.cc, subject, body_has: ['협력사 C 품질 담당자님', '2026-11-12까지 회신', name, '이품질 드림'], attach: name }));
  const out = cp.spawnSync('python3', ['-I', require('node:url').fileURLToPath(new URL('../scripts/check-mail.py', import.meta.url)), path.join(dir, 'm.eml'), path.join(dir, 'a.xlsx'), path.join(dir, 'w.json')], { encoding: 'utf8' });
  fs.rmSync(dir, { recursive: true, force: true });
  assert.equal(out.status, 0, out.stdout + out.stderr);
  console.log('     ' + out.stdout.trim());
});

const { versions } = await import('../scripts/bump-version.mjs');
test('판 맞추기 — index.html ?v= · logic.js · app.js 가 같은 판 (scripts/bump-version.mjs)', () => {
  const v = versions();
  assert.equal(v.logic, L.VERSION, 'logic.js VERSION'); assert.equal(v.app, L.VERSION, 'app.js APP_VERSION');
  assert.ok(v.html.length >= 9 && v.html.every(x => x === L.VERSION), 'index.html 의 모든 js·css 에 ?v=' + L.VERSION + ' — ' + JSON.stringify(v.html));
});

const single = await import('../scripts/build-single.mjs');
test('HTML 한 파일 — dist 가 지금 소스와 같음(낡지 않음) · 외부 파일 없음 · 같은 판', () => {
  const fs = require('node:fs');
  const { html, used } = single.buildSingle();
  assert.ok(fs.existsSync(single.OUT), 'dist 파일 없음 — node scripts/build-single.mjs');
  assert.ok(fs.readFileSync(single.OUT, 'utf8') === html, 'dist 가 낡았습니다 — node scripts/build-single.mjs (또는 bump-version) 로 다시 만드세요');
  assert.equal(used.length, 15);   // 9 + 로그인 키트 4(css/auth.css · vendor/supabase.js · js/auth-config.js · js/auth.js) + 게시판 2(css/board.css · js/board.js)
  assert.ok(!/<script src=|<link rel="stylesheet"/.test(html));
  assert.ok(html.includes("var VERSION = '" + L.VERSION + "'") && html.includes("var APP_VERSION = '" + L.VERSION + "'"));
  assert.ok(html.includes('window.__SINGLE_FILE__ = true') && html.includes("if(L.protocol==='file:')return;"));
  assert.ok(!/<\/script(?!>)/i.test(html.replace(/<\/script>/g, '')), '스크립트 안에 끊는 글자 없음');
});

console.log('\n' + passed + '개 통과' + (process.exitCode ? ' — 실패 있음' : ''));
