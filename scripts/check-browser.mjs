// 헤드리스 브라우저로 주요 흐름 점검 (선택 — playwright-core 가 있는 PC 에서만)
//   PLAYWRIGHT_CORE=/경로/node_modules/playwright-core node scripts/check-browser.mjs
//   SINGLE=1 … → dist 의 HTML 한 파일을 file:// 로 열고 http·https 요청을 모두 끊은 채(사내 PC 오프라인 흉내) 같은 흐름을 돕니다
//   WIDTHS=390,1280 … → 넘침 점검 폭을 줄임
// file:// 로 index.html 을 열어 불량 등록(직접 입력·엑셀 붙여넣기·수정·엑셀 백업) → 예시 데이터 → 불량 분석
// → Master Plan 초안 → 오른쪽 칸 실적 입력 → 계획 확정 → AI 답 반영 → 추진현황 → 보고서 내려받기(서식 확인) → 설정 순서로 누르고,
// JS 오류 · 외부 주소 요청(사내망 전용이라 0건이어야 함) · 폭별 가로 넘침을 잽니다.
import { createRequire } from 'node:module';
import path from 'node:path';
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
const require = createRequire(import.meta.url);
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const { chromium } = require(process.env.PLAYWRIGHT_CORE || 'playwright-core');
const XLSX = require(path.join(root, 'vendor/xlsx.full.min.js'));
const SINGLE = !!process.env.SINGLE;
const url = SINGLE ? 'file://' + path.join(root, 'dist', 'data10-07_협력사공정불량_개선관리.html') : 'file://' + path.join(root, 'index.html');
const H = 4; // MASTER PLAN 시트 머리 줄(0부터)

const errors = [];
const problems = [];
const browser = await chromium.launch();
const ctx = await browser.newContext({ acceptDownloads: true, viewport: { width: 1280, height: 900 } });
const page = await ctx.newPage();
page.on('pageerror', e => errors.push('pageerror: ' + e.message));
page.on('console', m => { if (m.type() === 'error') errors.push('console: ' + m.text()); });
page.on('dialog', d => d.accept());
const external = [], blocked = [];
page.on('request', r => { const u = r.url(); if (!/^(file|data|blob):/.test(u)) external.push(u); });
// HTML 한 파일 점검: 인터넷이 없는 사내 PC 처럼 http·https 를 모두 끊습니다(끊긴 요청은 따로 셈)
if (SINGLE) await ctx.route(/^https?:\/\//, route => { blocked.push(route.request().url()); return route.abort(); });

function ok(cond, msg) { if (!cond) problems.push(msg); else console.log('  ok  ' + msg); }
async function go(hash) { await page.evaluate(h => { location.hash = h; }, hash); await page.waitForTimeout(150); }
async function clickText(sel, text) { await page.locator(sel, { hasText: text }).first().click(); await page.waitForTimeout(150); }

await page.goto(url);
if (SINGLE) {
  ok(await page.evaluate(() => window.__SINGLE_FILE__ === true && document.querySelectorAll('script[src],link[rel=stylesheet]').length === 0), 'HTML 한 파일 — 외부 script·css 없음');
  ok(await page.evaluate(() => location.protocol === 'file:' && !location.search.includes('_fg')), 'fresh-guard 가 file:// 에서 아무것도 안 함(주소 그대로)');
}
// 점검기 자체 확인용: BREAK=1 이면 일부러 넓은 상자를 넣어 넘침을 잡는지 봅니다
if (process.env.BREAK) await page.addStyleTag({ content: '.page-head { min-width: 900px; }' });
// 점검기 자체 확인용: NET=1 이면 일부러 외부 주소를 불러 요청 점검이 잡는지 봅니다
if (process.env.NET) await page.evaluate(() => { const i = new Image(); i.src = 'https://example.com/x.png'; });

// ── 불량 등록 (파일 없이) ──
await go('#/entry');
const f = '#entryForm';
await page.fill(f + ' input[type=date]', '2026-10-05');
await page.fill(f + ' input[list="dl-sup"]', '협력사 가');
await page.fill(f + ' input[placeholder="예: TP-0001"]', 'TP-9001');
await page.fill(f + ' textarea >> nth=0', '브래킷 용접부 균열 발생');
ok((await page.textContent('#entryCard')).includes('자동 분류: 용접불량'), '내용 입력 중 자동 분류 안내(용접불량)');
await page.fill(f + ' input[type=number]', '2');
await page.fill(f + ' textarea >> nth=1', '전수 선별');
await clickText(f + ' button', '등록');
ok(await page.inputValue(f + ' input[list="dl-sup"]') === '협력사 가', '등록 후 협력사 남겨 둠');
await page.fill(f + ' textarea >> nth=0', '호스 체결부 조임 불량');
await page.selectOption(f + ' select', '미조임');
await clickText(f + ' button', '등록');
await page.fill(f + ' textarea >> nth=0', '');
await clickText(f + ' button', '등록');
ok((await page.textContent('#entryCard')).includes('내용(문제점)을 적어 주세요'), '필수 칸 빠지면 안내');
// 엑셀 붙여넣기 — 머리행 없음(순서대로) 2줄 + 유형 칸 「위치·치수」
await page.fill('textarea.paste-box', '2026-10-06\t협력사 나\tTP-9002\t커버\t위치·치수\t홀 위치 틀림\t1\t재작업\n2026.10.07\t협력사 나\tTP-9003\t\t\t오일 샘 발생\t\t');
await clickText('button', '표 읽기');
ok((await page.textContent('.ai-preview')).includes('읽은 결과 — 2건'), '붙여넣기 2건 읽음');
await clickText('button', '2건 넣기');
const listText = await page.textContent('main');
ok(listText.includes('3. 등록된 불량 4건'), '등록된 불량 4건');
ok(listText.includes('위치/치수') && listText.includes('누유/누수') && listText.includes('미조임'), '유형: 위치/치수(유형 칸) · 누유/누수(키워드) · 미조임(직접 고름)');
// 수정
await page.locator('main table.list tbody tr', { hasText: '오일 샘' }).locator('button', { hasText: '수정' }).click();
await page.waitForTimeout(150);
await page.fill(f + ' textarea >> nth=1', '협력사 반품');
await clickText(f + ' button', '수정 저장');
ok((await page.locator('main table.list tbody tr', { hasText: '오일 샘' }).textContent()).includes('협력사 반품'), '수정 저장(조치)');
// 엑셀 백업 → 다시 읽기
{
  const [d] = await Promise.all([page.waitForEvent('download'), clickText('button', '엑셀 백업 받기')]);
  const fp = path.join(root, 'samples', '.check-backup.xlsx'); await d.saveAs(fp);
  const b = XLSX.read(fs.readFileSync(fp), { type: 'buffer' }); fs.unlinkSync(fp);
  const rows = XLSX.utils.sheet_to_json(b.Sheets['불량 목록'], { header: 1, raw: false });
  ok(rows.length === 5 && rows[0].join(',') === '발생일자,협력사,품번,품명,불량유형,내용,수량,조치,관리번호,캠퍼스', '엑셀 백업 4건 + 머리행(관리번호 · 캠퍼스 포함)');
}
// 예시 데이터로 바꾸기(확인 창)
await go('#/upload');
await clickText('button', '예시 데이터로 시작');
await page.locator('#dialog button', { hasText: '바꾸기' }).click();
await page.waitForTimeout(200);
ok(await page.locator('#sampleBanner').isVisible(), '예시 데이터 띠 보임');
ok((await page.textContent('main')).includes('개선 필요 항목'), '불량 분석 화면');
// 4차 — 캠퍼스: 예시도 울산 · 인천 두 파일을 합친 것. 인천만 다시 올리면 인천 RAW 만 바뀌고 울산은 그대로
const kpiTotal = async () => Number((await page.locator('.dk', { hasText: '불량 발생 건수' }).first().locator('.dk-v').textContent()).replace(/[^0-9]/g, ''));
const allN = await kpiTotal();
ok((await page.textContent('main')).includes('울산캠퍼스') && (await page.textContent('main')).includes('인천캠퍼스'), '분석 KPI 에 캠퍼스별 건수');
await page.selectOption('main select[aria-label="캠퍼스"]', '인천');
await page.waitForTimeout(150);
const inchN = await kpiTotal();
await page.selectOption('main select[aria-label="캠퍼스"]', '울산');
await page.waitForTimeout(150);
const ulsN = await kpiTotal();
ok(inchN > 0 && ulsN > 0 && inchN + ulsN === allN, '캠퍼스 거르기 — 울산 ' + ulsN + ' + 인천 ' + inchN + ' = 전체 ' + allN);
await page.selectOption('main select[aria-label="캠퍼스"]', '');
await go('#/upload');
ok(await page.locator('#campusCard .campus-slot').count() === 2, '업로드 화면 — 울산 · 인천 캠퍼스 칸 2개');
async function campusUpload(i, file) {
  await go('#/upload');
  await page.locator('#campusCard .campus-slot').nth(i).locator('input[type=file]').setInputFiles(path.join(root, 'samples', file));
  await page.waitForSelector('#importCard');
  const t = await page.textContent('#importCard');
  await clickText('#importCard button', '파일 가져오기');
  await page.waitForTimeout(250);
  return t;
}
// 예시 데이터 위에 실제 파일을 올리면 예시를 비우고 시작 → 울산 · 인천을 차례로 올리고 → 인천만 한 번 더
await campusUpload(0, '예시_울산캠퍼스_RAW.xlsx');
ok(!(await page.locator('#sampleBanner').isVisible()), '실제 파일을 올리면 예시 데이터를 비움');
ok((await campusUpload(1, '예시_인천캠퍼스_RAW.xlsx')).includes('인천캠퍼스'), '인천 칸에 올린 파일 = 인천캠퍼스로 표시');
await go('#/analysis');
ok(await kpiTotal() === allN, '울산 + 인천 두 파일을 합친 건수 = 예시와 같음 ' + allN);
await campusUpload(1, '예시_인천캠퍼스_RAW.xlsx');
await go('#/upload');
const upText = await page.textContent('main');
ok(upText.includes('전에 올린 ' + inchN + '건을 바꿈') && upText.includes('울산캠퍼스 ' + ulsN) && upText.includes('인천캠퍼스 ' + inchN), '인천만 다시 올림 — 인천 ' + inchN + '건 바꿈, 울산 ' + ulsN + '건 그대로');
await go('#/analysis');

// 분류 수동 지정
const sel = page.locator('select.type-sel').first();
await sel.selectOption('기타');
await page.waitForTimeout(100);
ok(await page.locator('select.type-sel.manual').count() >= 1, '불량유형 직접 고르기 저장');

// Master Plan
await go('#/plan/' + encodeURIComponent('협력사 C'));
await clickText('button', '분석 및 초안 생성');
const nRows = await page.locator('table.gantt tbody tr.clickable').count();
ok(nRows > 5, 'Master Plan 초안 행 ' + nRows + '개');
ok((await page.textContent('table.gantt thead')).includes('2026년'), '간트 머리에 연·월');
ok(await page.locator('.dkpis.six .dk').count() === 6, 'KPI 6칸');
ok(!(await page.locator('#aiCard').getAttribute('open') !== null), 'AI 칸은 기본으로 접힘(선택 기능)');

// 오른쪽 칸 — 실적·완료
await page.locator('table.gantt tbody tr.clickable').first().click();
ok(await page.locator('#drawer').isVisible(), '행 누르면 오른쪽 칸');
await page.locator('#drawer input[placeholder^="예: 홍길동"]').fill('김담당 책임');
await page.locator('#drawer .st-btn.st-done').click();
await clickText('#drawer button', '저장');
ok(await page.locator('#drawer').isHidden(), '저장 후 닫힘');
ok(await page.locator('table.gantt .pill.st-done').count() === 1, '완료 상태 1건');
ok(await page.locator('table.gantt .g-act.st-done').count() === 1, '실적 막대(녹색)');

// 확정
await clickText('button', '계획 확정');
await page.locator('#dialog button', { hasText: '확정' }).last().click();
await page.waitForTimeout(150);
ok((await page.textContent('.gantt-card')).includes('확정 v1'), '계획 확정 v1');

// AI 답 반영(선택 기능 — 칸을 펼쳐서)
await page.locator('#aiCard > summary').click();
ok((await page.textContent('#copyPrompt')).includes('Copilot'), '프롬프트 복사 단추 = Copilot 에 붙여 넣기(기본)');
await page.locator('.ai-targets button', { hasText: 'Gemini' }).click();
await page.waitForTimeout(100);
ok((await page.textContent('#copyPrompt')).includes('Gemini') && (await page.textContent('#aiCard')).includes('임시 채팅'), 'Gemini(임시 채팅) 고르면 안내가 바뀜');
await page.locator('.ai-targets button', { hasText: 'Copilot' }).click();
await page.waitForTimeout(100);
const prompt = await page.locator('textarea.prompt-box[readonly]').inputValue();
ok(!/협력사 [A-F]|TP-\d{4}/.test(prompt), '프롬프트에 협력사·품번 없음(별칭)');
const alias = (/협력사: (업체\d+)/.exec(prompt) || [])[1];
await page.locator('textarea.prompt-box.answer').fill(['요약: ' + alias + ' 용접 균열이 늘었어.', '| 불량유형 | 개선대책 | 담당부서 | 시작 | 기간 |', '|---|---|---|---|---|', '| 용접 불량 | 예열 온도 기록지 도입 | 품질 | 0 | 2 |', '| 위치·치수 | 지그 점검표 개정 | 생산기술팀 | 1 | 3 |', '확인 필요: 예열 설비 유무'].join('\n'));
await clickText('button', '답 읽기');
ok((await page.textContent('.ai-preview')).includes('개선대책 2건'), 'AI 답 2건 읽음');
await clickText('button', '계획에 반영');
ok((await page.textContent('.ai-sum')).includes('협력사 C 용접 균열이 늘었어'), 'AI 요약 별칭 되돌림');
ok((await page.textContent('table.gantt')).includes('위치/치수'), 'AI 답의 「위치·치수」 → 위치/치수');
ok((await page.textContent('.gantt-card')).includes('이후 계획 수정 있음'), '확정 뒤 수정 표시');

// 개선활동 추가
await clickText('button', '개선활동 추가');
await page.locator('#drawer textarea').first().fill('협력사 품질회의 정례화');
await page.locator('#drawer .st-btn.st-hold').click();
await clickText('#drawer button', '저장');
ok(await page.locator('table.gantt .pill.st-hold').count() === 1, '보류 활동 추가');

// 추진현황·Excel 추출(두 시트)
await go('#/progress');
ok(await page.locator('table.gantt tbody tr.clickable').count() > 0, '추진현황 간트');
ok(await page.locator('.ptiles .pt').count() === 5 && (await page.textContent('.ptiles')).includes('진행중') && (await page.textContent('.ptiles')).includes('미완료') && !(await page.textContent('main')).includes('예정 '), '현황 타일 전체·진행중·완료·미완료·보류');
await go('#/report');
await page.selectOption('select[aria-label="보고서 협력사"]', '협력사 C');
await page.waitForTimeout(150);
const [dl] = await Promise.all([page.waitForEvent('download'), clickText('button', 'Excel 추출')]);
const file = path.join(root, 'samples', '.check-report.xlsx');
await dl.saveAs(file);
const raw = fs.readFileSync(file);
const wb = XLSX.read(raw, { type: 'buffer', cellNF: true });
if (process.env.KEEP) fs.copyFileSync(file, process.env.KEEP);
fs.unlinkSync(file);
ok(wb.SheetNames.join(',') === '불량목록,MASTER PLAN', '엑셀 시트 2개: ' + wb.SheetNames.join(','));
const mp = XLSX.utils.sheet_to_json(wb.Sheets['MASTER PLAN'], { header: 1, raw: true, defval: null });
ok(mp[H][0] === 'ID' && mp[H][11] === '진행상태' && !mp[H].some(x => /근거|진행률/.test(String(x))), 'MASTER PLAN 머리 — ID·진행상태, 근거·진행률 칸 없음');
ok(mp.some(r => r[4] === '협력사 품질회의 정례화' && r[11] === '보류'), 'MASTER PLAN 에 보류 활동');
ok(mp.slice(H + 1).every(r => !r[0] || ['진행중', '완료', '미완료', '보류'].includes(r[11])), 'MASTER PLAN 진행상태는 네 값만');
ok(wb.Sheets['불량목록'].C2.t === 'n' && wb.Sheets['불량목록'].C2.z === 'yyyy-mm-dd' && wb.Sheets['불량목록'].B1.v === '캠퍼스', '불량목록 — 캠퍼스 칸 · 발생일자 = 엑셀 날짜');
{
  const z = XLSX.CFB.read(raw, { type: 'buffer' });
  const s2 = Buffer.from(z.FileIndex[z.FullPaths.indexOf('Root Entry/xl/worksheets/sheet2.xml')].content).toString('utf8');
  ok((s2.match(/<cfRule type="cellIs"/g) || []).length === 4 && /<conditionalFormatting sqref="L\d+:L\d+">/.test(s2), 'MASTER PLAN 진행상태 칸 = 조건부 서식 4개(엑셀에서 값을 바꾸면 색이 바뀜)');
}

// 협력사가 고쳐 회신한 것처럼 — 담당자·진행률(1줄), 완료+실적(1줄), 완료인데 실적 완료일 없음(확인 필요 1줄), 새 활동 1줄
{
  const ws = wb.Sheets['MASTER PLAN'];
  const set = (r, c, v) => { ws[XLSX.utils.encode_cell({ r, c })] = typeof v === 'number' ? { t: 'n', v } : { t: 's', v }; };
  // 이 흐름의 협력사 C 계획 = AI 답 2줄(미완료) + 보류 1줄
  const open = [], hold = []; for (let r = H + 1; r < mp.length; r++) if (mp[r][0]) (mp[r][11] === '보류' ? hold : mp[r][11] !== '완료' ? open : []).push(r);
  if (open.length < 2 || !hold.length) problems.push('회신 점검용 행이 모자람: 미완료 ' + open.length + ' · 보류 ' + hold.length);
  set(open[0], 6, '협력사 김담당');
  set(open[1], 9, '2026-10-20'); set(open[1], 10, '2026-10-26'); set(open[1], 11, '완료');
  set(hold[0], 11, '완료');   // 실적 완료일 없이 완료 → 확인 필요
  const nr = mp.length + 1;
  ['', '', '협력사 C', '용접불량', '용접 지그 교체(협력사 추가)', '생산기술', '박생기', '2026-11-16', '2026-11-29', '', '', '진행중', ''].forEach((v, c) => { if (v !== '') set(nr, c, v); });
  ws['!ref'] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: nr, c: 13 } });
  fs.writeFileSync(file, XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' }));
}
await go('#/reply');
await page.setInputFiles('#replyFile', file);
await page.waitForSelector('#replyPreview');
const pv = await page.textContent('#replyPreview');
ok(/바뀐 행\s*2/.test(pv) && /새 활동\s*1/.test(pv) && /확인 필요\s*1/.test(pv), '회신 미리 보기 — 바뀐 행 2 · 새 활동 1 · 확인 필요 1');
ok(pv.includes('협력사 김담당'), '바뀐 칸(담당자) 보임');
await page.fill('#replyPreview input[aria-label="반영한 사람"]', '품질 이준호');
await page.locator('#replyApply').click();
await page.waitForTimeout(200);
ok((await page.textContent('#replyHistory')).includes('품질 이준호') && (await page.textContent('#replyHistory')).includes('.check-report.xlsx'), '반영 이력(누가·파일)');
await page.setInputFiles('#replyFile', file);
await page.waitForSelector('#replyPreview');
ok((await page.textContent('#replyApply')).includes('반영할 것 없음') && (await page.textContent('#replyPreview')).includes('이미 반영한 파일'), '같은 파일 다시 올리면 반영할 것 없음(멱등)');
fs.unlinkSync(file);
await go('#/plan/' + encodeURIComponent('협력사 C'));
ok((await page.textContent('table.gantt')).includes('용접 지그 교체(협력사 추가)'), '회신의 새 활동이 간트에');
ok(await page.locator('table.gantt .pill.st-done').count() === 1 && await page.locator('table.gantt .pill.st-hold').count() === 1, '회신으로 완료 1건 · 확인 필요 줄(보류)은 그대로');
// 대시보드 현황 타일
await go('#/home');
ok((await page.textContent('.ptiles')).includes('전체 과제'), '대시보드 개선 MASTER PLAN 현황');
// 저장 유지 — 다시 열어도 데이터가 남음(file:// 의 localStorage)
await page.reload(); await page.waitForTimeout(300);
await go('#/reply');
ok((await page.textContent('#replyHistory')).includes('반영 이력 1건'), '다시 열어도 데이터·이력이 남음(localStorage)');

// 설정·대시보드
await go('#/settings');
ok(await page.locator('.dict-item').count() === 10, '불량유형 사전 10종');
ok((await page.textContent('#deptCard')).includes('품질 · 생산기술 · 영업 · 설계 · 구매'), '담당부서 기본 5종 안내');
ok(await page.inputValue('#deptCard textarea') === '품질, 생산기술, 영업, 설계, 구매', '담당부서 기본값');
ok(SINGLE ? (await page.textContent('#singleCard')).includes('HTML 한 파일 판입니다') : (await page.getAttribute('#singleDownload', 'href')) === 'dist/data10-07_협력사공정불량_개선관리.html', 'HTML 파일로 받기 단추(한 파일 판에서는 안내)');
ok((await page.textContent('main')).includes('사내망에서 쓰기'), '사내망 안내');
// 4차 — 아웃룩 메일: 설정에 수신 · 참조 → 협력사 C 「메일 만들기」 → .eml 내려받기 → 파이썬 email 패키지로 다시 읽기
await page.fill('#mailCard input[aria-label="수신1 이름"]', '김품질');
await page.fill('#mailCard input[aria-label="수신1 EMAIL"]', 'q1@example.com');
await page.fill('#mailCard input[aria-label="참조1 이름"]', 'Lead');
await page.fill('#mailCard input[aria-label="참조1 EMAIL"]', 'lead@example.com');
await page.fill('#mailCard input[aria-label="보내는 사람 이름"]', '이품질');
await page.locator('#mailSave').click();
await page.waitForTimeout(150);
await go('#/plan/' + encodeURIComponent('협력사 C'));
await clickText('button', '메일 만들기');
const dlg = await page.textContent('#dialog');
ok(dlg.includes('브라우저는 보안상 아웃룩을 직접') && dlg.includes('q1@example.com') && dlg.includes('lead@example.com'), '메일 만들기 창 — 안내 · 받는 사람 · 참조');
ok(await page.locator('#dialog button', { hasText: '메일 프로그램 열기(첨부 없음)' }).count() === 1, 'mailto(첨부 없음) 단추');
{
  const subj = await page.inputValue('#dialog input[aria-label="메일 제목"]');
  ok(subj.includes('협력사 C') && !/\{[^}]+\}/.test(subj + await page.inputValue('#dialog textarea[aria-label="메일 본문"]')), '제목 · 본문 자리표시가 채워짐: ' + subj);
  const [d] = await Promise.all([page.waitForEvent('download'), page.locator('#emlBtn').click()]);
  const eml = path.join(root, 'samples', '.check-mail.eml'); await d.saveAs(eml);
  const txt = fs.readFileSync(eml, 'latin1');
  const bnd = /boundary="([^"]+)"/.exec(txt)[1];
  const att = txt.split('--' + bnd).find(p => /Content-Disposition: attachment/.test(p));
  const xlsx = path.join(root, 'samples', '.check-mail.xlsx'), want = path.join(root, 'samples', '.check-mail.json');
  fs.writeFileSync(xlsx, Buffer.from(att.split('\r\n\r\n')[1].replace(/\s+/g, ''), 'base64'));
  const fname = decodeURIComponent(/filename\*=UTF-8''([^\r\n]+)/.exec(att)[1]);
  fs.writeFileSync(want, JSON.stringify({ to: [{ name: '김품질', email: 'q1@example.com' }], cc: [{ name: 'Lead', email: 'lead@example.com' }], subject: subj, body_has: ['협력사 C', '이품질 드림'], attach: fname }));
  const cp = require('node:child_process');
  const r = cp.spawnSync('python3', ['-I', path.join(root, 'scripts', 'check-mail.py'), eml, xlsx, want], { encoding: 'utf8' });
  ok(r.status === 0, '.eml 파이썬으로 다시 읽기 — ' + (r.stdout + r.stderr).trim());
  ok(d.suggestedFilename().endsWith('.eml') && fname.startsWith('MasterPlan_협력사C'), '내려받은 이름 ' + d.suggestedFilename() + ' · 첨부 ' + fname);
  for (const f of [eml, xlsx, want]) fs.unlinkSync(f);
}
await go('#/settings');
await go('#/home');
ok((await page.textContent('main')).includes('협력사별 추진현황'), '대시보드');

// 폭별 가로 넘침 — 안쪽 가로 스크롤 상자(overflow-x auto) 안은 의도된 것이라 뺍니다
const WIDTHS = process.env.WIDTHS ? process.env.WIDTHS.split(',').map(Number) : [360, 390, 768, 1024, 1280, 1440];
for (const w of WIDTHS) {
  await page.setViewportSize({ width: w, height: 900 });
  for (const r of ['#/home', '#/entry', '#/upload', '#/analysis', '#/plan', '#/progress', '#/report', '#/reply', '#/settings']) {
    await go(r);
    const bad = await page.evaluate(() => {
      const cw = document.documentElement.clientWidth, out = [];
      for (const el of document.querySelectorAll('main *, .topbar *')) {
        const b = el.getBoundingClientRect();
        if (!b.width || !b.height) continue;
        let p = el.parentElement, inner = false;
        while (p) { const ox = getComputedStyle(p).overflowX; if (ox === 'auto' || ox === 'scroll' || ox === 'hidden') { inner = true; break; } p = p.parentElement; }
        if (!inner && b.right > cw + 1) out.push(el.tagName + '.' + el.className + ' ' + Math.round(b.right) + '>' + cw);
      }
      return { out: out.slice(0, 3), page: document.documentElement.scrollWidth > cw + 1 };
    });
    if (bad.page || bad.out.length) problems.push(w + 'px ' + r + ' 넘침: ' + bad.out.join(' | '));
  }
}
console.log('  폭 ' + WIDTHS.join('·') + ' × 화면 9개 넘침 점검 끝');
if (external.length) problems.push('외부 주소 요청 ' + external.length + '건: ' + [...new Set(external)].slice(0, 3).join(', '));
else console.log('  ok  외부 주소 요청 0건(사내망 전용)' + (SINGLE ? ' · 끊긴 요청 ' + blocked.length + '건' : ''));

await browser.close();
if (errors.length) problems.push(...errors);
if (problems.length) { console.error('\n문제 ' + problems.length + '건\n' + problems.join('\n')); process.exitCode = 1; }
else console.log('\n브라우저 점검 통과');
