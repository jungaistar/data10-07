/* Excel(.xlsx) 쓰기 — 웹 화면의 표 모양 그대로 내보내기 위한 서식 층.
   vendor 의 SheetJS(커뮤니티판)는 글꼴 굵게·칸 색을 쓰지 못해서, SheetJS 가 만든 xlsx 를 다시 열어
   styles.xml 에 서식(굵은 머리글·칸 색·테두리)을 더하고 해당 칸에 s="번호" 를 붙입니다.
   외부 서버·CDN 을 부르지 않습니다(사내망 전용). node 에서도 불러 시험합니다.

   build(XLSX, sheets, meta) → Uint8Array
     sheets = { 시트이름: [[칸, …], …] }  — 「2026-10-05」 꼴 글자는 엑셀 날짜(yyyy-mm-dd)로 바꿉니다
     meta   = { 시트이름: { head:[행번호…], title:[행번호…], section:[행번호…], cells:[{r,c,style}],
                            merges:[{s:{r,c},e:{r,c}}], widths:[글자수…], freeze:{rows,cols}, filter:'A1:K20',
                            hidden:[열번호…](숨김 열), lists:[{sqref:'L6:L500', items:['완료',…], soft}](고르기 목록 — soft 면 다른 값도 받음),
                            cf:[{sqref:'L6:L500', rules:[{equal:'완료', style:'done'}, …]}](조건부 서식 — 칸 값이 그 글자와 같으면 그 색) } }
     style  = head(굵게·회색 바탕·테두리) · edit(협력사가 고치는 머리 칸, 노랑) · title(굵게·크게) · section(굵게) · plan(계획 칸)
              · done · open · hold · doing · late · planned(진행상태 색) · today
     조건부 서식(4차): SheetJS 커뮤니티판은 조건부 서식을 쓰지 못해 styles.xml 에 <dxfs>(차등 서식)를 더하고 시트 XML 에
     <conditionalFormatting> 을 직접 넣습니다. 칸에 박아 넣는 색과 달리 협력사가 진행상태를 바꾸면 엑셀이 색을 바로 바꿉니다.
     「미완료」에 「완료」가 들어 있어 「포함」 규칙이 아니라 「같음(cellIs equal)」 규칙을 씁니다. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.SQXlsx = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  var DATE_RE = /^\d{4}-\d\d-\d\d$/;
  function serial(s) { return (Date.UTC(+s.slice(0, 4), +s.slice(5, 7) - 1, +s.slice(8, 10)) - Date.UTC(1899, 11, 30)) / 86400000; }
  /* 글자 폭 — 한글 등 넓은 글자는 2칸으로 셉니다 */
  function textWidth(v) {
    var s = v == null ? '' : String(v), w = 0;
    for (var i = 0; i < s.length; i++) w += s.charCodeAt(i) > 0x2e80 ? 2 : 1;
    return w;
  }
  function autoWidths(aoa, given) {
    var w = [];
    aoa.forEach(function (row) {
      (row || []).forEach(function (v, c) {
        var t = DATE_RE.test(String(v)) ? 11 : textWidth(v) + 2;
        w[c] = Math.max(w[c] || 6, Math.min(t, 50));
      });
    });
    return w.map(function (x, c) { return { wch: given && given[c] ? given[c] : x }; });
  }
  function colName(c) { var s = ''; c++; while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; }
  function ref(r, c) { return colName(c) + (r + 1); }

  /* 더할 서식 — 칸 색은 화면 색(css)과 같은 계열 */
  /* 진행상태 색(4차) — 완료 초록 · 진행중 주황 · 미완료 빨강 · 보류 회색 (화면 css 와 같은 색) */
  var FILL = { head: 'FFE6EBF2', edit: 'FFFFE699', plan: 'FFD6E4F5', done: 'FF1F9D55', open: 'FFE8830C', hold: 'FF7F8C8D', doing: 'FFE8830C', late: 'FFC0392B', planned: 'FF6B7785', today: 'FFFFF4CC' };
  var CF_KEYS = ['done', 'doing', 'late', 'hold'];
  var STYLES = ['head', 'title', 'section', 'plan', 'done', 'open', 'hold', 'today', 'edit', 'doing', 'late', 'planned'];

  function addStyles(xml) {
    var nFont = +(/<fonts count="(\d+)"/.exec(xml) || [])[1];
    var nFill = +(/<fills count="(\d+)"/.exec(xml) || [])[1];
    var nBorder = +(/<borders count="(\d+)"/.exec(xml) || [])[1];
    var nXf = +(/<cellXfs count="(\d+)"/.exec(xml) || [])[1];
    if (!(nFont >= 1 && nFill >= 2 && nBorder >= 1 && nXf >= 1)) throw new Error('styles.xml 모양이 예상과 다릅니다');
    var fonts = '<font><b/><sz val="11"/><color theme="1"/><name val="맑은 고딕"/><family val="2"/></font>' +   // nFont   : 굵게
      '<font><b/><sz val="14"/><color theme="1"/><name val="맑은 고딕"/><family val="2"/></font>' +             // nFont+1 : 제목
      '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="맑은 고딕"/><family val="2"/></font>';         // nFont+2 : 흰 글자(진한 칸)
    var fillKeys = ['head', 'plan', 'done', 'open', 'hold', 'today', 'edit', 'doing', 'late', 'planned'], fillId = {};
    var fills = fillKeys.map(function (k, i) { fillId[k] = nFill + i; return '<fill><patternFill patternType="solid"><fgColor rgb="' + FILL[k] + '"/><bgColor indexed="64"/></patternFill></fill>'; }).join('');
    var thin = '<border><left style="thin"><color rgb="FFB7C3D3"/></left><right style="thin"><color rgb="FFB7C3D3"/></right><top style="thin"><color rgb="FFB7C3D3"/></top><bottom style="thin"><color rgb="FFB7C3D3"/></bottom><diagonal/></border>';
    var B = nBorder, center = '<alignment horizontal="center" vertical="center"/>', wrap = '<alignment vertical="center" wrapText="1"/>';
    function xf(font, fill, border, align) {
      return '<xf numFmtId="0" fontId="' + font + '" fillId="' + fill + '" borderId="' + border + '" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">' + align + '</xf>';
    }
    var defs = {
      head: xf(nFont, fillId.head, B, wrap),
      title: xf(nFont + 1, 0, 0, '<alignment vertical="center"/>'),
      section: xf(nFont, 0, 0, '<alignment vertical="center"/>'),
      plan: xf(0, fillId.plan, B, center),
      done: xf(nFont + 2, fillId.done, B, center),
      open: xf(nFont + 2, fillId.open, B, center),
      hold: xf(nFont + 2, fillId.hold, B, center),
      today: xf(nFont, fillId.today, B, center),
      edit: xf(nFont, fillId.edit, B, wrap),
      doing: xf(nFont + 2, fillId.doing, B, center),
      late: xf(nFont + 2, fillId.late, B, center),
      planned: xf(nFont + 2, fillId.planned, B, center)
    };
    var ids = {};
    var xfs = STYLES.map(function (k, i) { ids[k] = nXf + i; return defs[k]; }).join('');
    // 조건부 서식용 차등 서식(dxf) — 흰 굵은 글자 + 상태 색 바탕
    var nDxf = +(/<dxfs count="(\d+)"/.exec(xml) || [0, 0])[1], dxfIds = {};
    var dxfs = CF_KEYS.map(function (k, i) { dxfIds[k] = nDxf + i; return '<dxf><font><b/><color rgb="FFFFFFFF"/></font><fill><patternFill patternType="solid"><fgColor rgb="' + FILL[k] + '"/><bgColor rgb="' + FILL[k] + '"/></patternFill></fill></dxf>'; }).join('');
    if (/<dxfs count="\d+"\/>/.test(xml)) xml = xml.replace(/<dxfs count="\d+"\/>/, '<dxfs count="' + (nDxf + CF_KEYS.length) + '">' + dxfs + '</dxfs>');
    else if (/<dxfs count="\d+">/.test(xml)) xml = xml.replace(/<dxfs count="\d+">([\s\S]*?)<\/dxfs>/, function (m, inner) { return '<dxfs count="' + (nDxf + CF_KEYS.length) + '">' + inner + dxfs + '</dxfs>'; });
    else xml = xml.replace(/<tableStyles|<\/styleSheet>/, function (m) { return '<dxfs count="' + CF_KEYS.length + '">' + dxfs + '</dxfs>' + m; });
    xml = xml.replace(/<fonts count="\d+">([\s\S]*?)<\/fonts>/, function (m, inner) { return '<fonts count="' + (nFont + 3) + '">' + inner + fonts + '</fonts>'; })
      .replace(/<fills count="\d+">([\s\S]*?)<\/fills>/, function (m, inner) { return '<fills count="' + (nFill + fillKeys.length) + '">' + inner + fills + '</fills>'; })
      .replace(/<borders count="\d+">([\s\S]*?)<\/borders>/, function (m, inner) { return '<borders count="' + (nBorder + 1) + '">' + inner + thin + '</borders>'; })
      .replace(/<cellXfs count="\d+">([\s\S]*?)<\/cellXfs>/, function (m, inner) { return '<cellXfs count="' + (nXf + STYLES.length) + '">' + inner + xfs + '</cellXfs>'; });
    return { xml: xml, ids: ids, dxf: dxfIds };
  }
  /* 칸에 서식 번호 붙이기 · 틀 고정 */
  function escXml(v) { return String(v).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
  /* 고르기 목록(데이터 유효성) — OOXML 순서상 conditionalFormatting 뒤, hyperlinks·printOptions·pageMargins 앞에 넣습니다 */
  /* 조건부 서식(conditionalFormatting)은 OOXML 순서상 dataValidations 바로 앞이라, 고르기 목록과 함께 같은 자리에 넣습니다 */
  function addLists(xml, lists, cf, dxf) {
    var prio = 0;
    var cfx = (cf || []).map(function (b) {
      return '<conditionalFormatting sqref="' + b.sqref + '">' + b.rules.filter(function (r) { return dxf && dxf[r.style] != null; }).map(function (r) {
        return '<cfRule type="cellIs" dxfId="' + dxf[r.style] + '" priority="' + (++prio) + '" operator="equal"><formula>"' + escXml(String(r.equal).replace(/"/g, '""')) + '"</formula></cfRule>';
      }).join('') + '</conditionalFormatting>';
    }).join('');
    if ((!lists || !lists.length) && !cfx) return xml;
    var dv = cfx + (lists && lists.length ? '<dataValidations count="' + lists.length + '">' + lists.map(function (l) {
      return '<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="' + (l.soft ? '0' : '1') + '"' + (l.soft ? ' errorStyle="information"' : '') + ' sqref="' + l.sqref + '"><formula1>"' + escXml(l.items.join(',')) + '"</formula1></dataValidation>';
    }).join('') + '</dataValidations>' : '');
    var m = /<(hyperlinks|printOptions|pageMargins|pageSetup|headerFooter|rowBreaks|colBreaks|ignoredErrors|drawing|legacyDrawing|tableParts|extLst)[ >\/]/.exec(xml);
    var at = m ? m.index : xml.lastIndexOf('</worksheet>');
    return xml.slice(0, at) + dv + xml.slice(at);
  }
  function styleSheetXml(xml, targets, ids, freeze) {
    Object.keys(targets).forEach(function (r) {
      var id = ids[targets[r]];
      xml = xml.replace(new RegExp('<c r="' + r + '"(?: s="\\d+")?(?=[ >/])'), '<c r="' + r + '" s="' + id + '"');
    });
    if (freeze && (freeze.rows || freeze.cols)) {
      var top = ref(freeze.rows || 0, freeze.cols || 0);
      var pane = '<pane' + (freeze.cols ? ' xSplit="' + freeze.cols + '"' : '') + (freeze.rows ? ' ySplit="' + freeze.rows + '"' : '') + ' topLeftCell="' + top + '" activePane="' + (freeze.rows && freeze.cols ? 'bottomRight' : freeze.rows ? 'bottomLeft' : 'topRight') + '" state="frozen"/>';
      xml = xml.replace('<sheetView workbookViewId="0"/>', '<sheetView workbookViewId="0">' + pane + '</sheetView>');
    }
    return xml;
  }

  function build(XLSX, sheets, meta) {
    meta = meta || {};
    var wb = XLSX.utils.book_new(), names = [], plans = [];
    Object.keys(sheets).forEach(function (n) {
      var aoa = sheets[n], m = meta[n] || {};
      var ws = XLSX.utils.aoa_to_sheet(aoa.map(function (row) {
        return (row || []).map(function (v) { return typeof v === 'string' && DATE_RE.test(v) ? { t: 'n', v: serial(v), z: 'yyyy-mm-dd' } : v; });
      }));
      ws['!cols'] = autoWidths(aoa, m.widths);
      (m.hidden || []).forEach(function (c) { while (ws['!cols'].length <= c) ws['!cols'].push({ wch: 8 }); ws['!cols'][c].hidden = true; });
      if (m.merges && m.merges.length) ws['!merges'] = m.merges;
      if (m.filter) ws['!autofilter'] = { ref: m.filter };
      var safe = n.replace(/[\\\/?*\[\]:]/g, '-').slice(0, 31);
      XLSX.utils.book_append_sheet(wb, ws, safe);
      // 서식을 붙일 칸 목록
      var t = {};
      function rowAll(r, style) { var row = aoa[r] || []; for (var c = 0; c < row.length; c++) t[ref(r, c)] = style; }
      (m.title || []).forEach(function (r) { t[ref(r, 0)] = 'title'; });
      (m.section || []).forEach(function (r) { t[ref(r, 0)] = 'section'; });
      (m.head || []).forEach(function (r) { rowAll(r, 'head'); });
      (m.cells || []).forEach(function (x) { if (aoa[x.r] && x.c < aoa[x.r].length) t[ref(x.r, x.c)] = x.style; });
      names.push(safe); plans.push({ targets: t, freeze: m.freeze, lists: m.lists, cf: m.cf });
    });
    var out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    var CFB = XLSX.CFB, z = CFB.read(new Uint8Array(out), { type: 'array' });
    function file(p) { var i = z.FullPaths.indexOf('Root Entry/' + p); return i < 0 ? null : z.FileIndex[i]; }
    function text(f) { return utf8Decode(f.content); }
    var st = file('xl/styles.xml');
    var added = addStyles(text(st));
    st.content = utf8Encode(added.xml); st.size = st.content.length;
    plans.forEach(function (p, i) {
      var f = file('xl/worksheets/sheet' + (i + 1) + '.xml');
      if (!f) return;
      f.content = utf8Encode(addLists(styleSheetXml(text(f), p.targets, added.ids, p.freeze), p.lists, p.cf, added.dxf)); f.size = f.content.length;
    });
    var res = CFB.write(z, { fileType: 'zip', type: 'array' });
    return res instanceof Uint8Array ? res : new Uint8Array(res);
  }
  function utf8Decode(u8) {
    if (typeof TextDecoder !== 'undefined') return new TextDecoder('utf-8').decode(u8 instanceof Uint8Array ? u8 : new Uint8Array(u8));
    return Buffer.from(u8).toString('utf8');
  }
  function utf8Encode(s) {
    if (typeof TextEncoder !== 'undefined') return new TextEncoder().encode(s);
    return new Uint8Array(Buffer.from(s, 'utf8'));
  }

  return { build: build, serial: serial, autoWidths: autoWidths, colName: colName, ref: ref, STYLES: STYLES };
});
