// HTML 한 파일 만들기 — 사내 보안망이라 GitHub 에 들어가기 어려워(2026-10-27 수강생 답) 사내 PC 에 저장해 두고 더블클릭해 쓰는 판.
//   node scripts/build-single.mjs           → dist/data10-07_협력사공정불량_개선관리.html 을 새로 씀
//   node scripts/build-single.mjs --check   → 지금 소스로 만든 결과와 dist 파일이 같은지만 봄(다르면 종료 코드 1)
// index.html 이 부르는 css/js/vendor 를 모두 그 자리에 넣습니다(<style>·<script>). 외부 주소를 부르지 않고 file:// 로 열어도 돕니다.
// 같은 소스면 늘 같은 결과(날짜·시각을 넣지 않음) — 테스트가 dist 가 낡았는지 이 방법으로 잡습니다.
import fs from 'node:fs'; import path from 'node:path'; import { fileURLToPath } from 'node:url';
const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const SINGLE_NAME = 'data10-07_협력사공정불량_개선관리.html';
export const OUT = path.join(ROOT, 'dist', SINGLE_NAME);

/* <script> 안에 넣을 때 HTML 해석기가 끊지 않게: 「</script」 → 「<\/script」, 「<!--」 → 「<\!--」 (JS 에서는 같은 뜻) */
function safeScript(js) { return js.replace(/<\/(script)/gi, '<\\/$1').replace(/<!--/g, '<\\!--'); }
function safeStyle(css) { return css.replace(/<\/(style)/gi, '<\\/$1'); }
function local(ref) { return path.join(ROOT, ref.replace(/\?.*$/, '')); }

export function buildSingle() {
  let html = fs.readFileSync(path.join(ROOT, 'index.html'), 'utf8');
  html = html.replace(/\n?[ \t]*<!-- seo-kit[^>]*-->[\s\S]*?<!-- \/seo-kit -->/g, '');   // 공유 미리보기 메타 · 아이콘은 한 파일 판에 필요 없다(외부 주소)
  const used = [];
  html = html.replace(/<link rel="stylesheet" href="([^"]+)">/g, (m, href) => {
    if (/^(https?:)?\/\//.test(href)) throw new Error('외부 CSS: ' + href);
    used.push(href);
    return '<style>/* ' + href.replace(/\?.*$/, '') + ' */\n' + safeStyle(fs.readFileSync(local(href), 'utf8')) + '\n</style>';
  });
  let first = true;
  html = html.replace(/<script src="([^"]+)"><\/script>/g, (m, src) => {
    if (/^(https?:)?\/\//.test(src)) throw new Error('외부 스크립트: ' + src);
    used.push(src);
    const flag = first ? '<script>window.__SINGLE_FILE__ = true;</script>\n  ' : '';
    first = false;
    return flag + '<script>/* ' + src.replace(/\?.*$/, '') + ' */\n' + safeScript(fs.readFileSync(local(src), 'utf8')) + '\n</script>';
  });
  html = html.replace(/<title>([^<]*)<\/title>/, '<title>$1 (HTML 한 파일)</title>')
    .replace('<!doctype html>', '<!doctype html>\n<!-- data10-07 HTML 한 파일 판 — scripts/build-single.mjs 가 index.html · css · js · vendor(SheetJS) 를 합쳐 만든 파일입니다. 직접 고치지 말고 소스를 고친 뒤 다시 만드세요. -->');
  if (/<(script|link)[^>]+(src|href)="(?!data:|#)[^"]+"/.test(html)) throw new Error('합치지 못한 외부 파일이 남았습니다');
  return { html, used };
}

if (process.argv[1] && fileURLToPath(import.meta.url) === path.resolve(process.argv[1])) {
  const { html, used } = buildSingle();
  if (process.argv[2] === '--check') {
    const same = fs.existsSync(OUT) && fs.readFileSync(OUT, 'utf8') === html;
    console.log(same ? 'dist 같음: ' + path.relative(ROOT, OUT) : 'dist 가 지금 소스와 다릅니다 — node scripts/build-single.mjs 로 다시 만드세요');
    process.exitCode = same ? 0 : 1;
  } else {
    fs.mkdirSync(path.dirname(OUT), { recursive: true });
    fs.writeFileSync(OUT, html);
    console.log('만듦: ' + path.relative(ROOT, OUT) + ' (' + (Buffer.byteLength(html) / 1024).toFixed(0) + ' KB, 합친 파일 ' + used.length + '개)');
  }
}
