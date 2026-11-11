#!/usr/bin/env python3
"""메일 초안(.eml)·MASTER PLAN 엑셀 검사 — 브라우저 밖(파이썬 표준 email 패키지 · openpyxl)에서 다시 읽어 봅니다.

실행: python3 -I scripts/check-mail.py <메일.eml> <첨부와 같아야 할.xlsx> <기대값.json>
기대값 = {"to": ["이름 <a@b>", …], "cc": […], "subject": "…", "body_has": ["…"], "attach": "파일이름.xlsx"}
- 머리글: RFC 2047 로 적은 한글 제목·이름이 그대로 풀리는지, X-Unsent: 1, multipart/mixed
- 첨부: base64 를 풀면 원본 엑셀과 바이트가 똑같은지, 파일 이름(RFC 2231)이 한글 그대로인지
- 엑셀: openpyxl 로 MASTER PLAN 시트의 진행상태 칸 조건부 서식 4개(완료 초록 · 진행중 주황 · 미완료 빨강 · 보류 회색)가 읽히는지
test/logic.test.mjs 가 python3 가 있으면 부릅니다. 실패하면 0 이 아닌 값으로 끝납니다."""
import email, json, sys, io
from email import policy
from email.headerregistry import Address

def main(eml_path, xlsx_path, want_path):
    raw = open(eml_path, 'rb').read()
    want = json.load(open(want_path, encoding='utf-8'))
    msg = email.message_from_bytes(raw, policy=policy.default)
    errs = []
    def ok(cond, what):
        if not cond: errs.append(what)
    ok(msg.get_content_type() == 'multipart/mixed', 'multipart/mixed 아님: ' + msg.get_content_type())
    ok(msg['X-Unsent'] == '1', 'X-Unsent: 1 없음')
    ok(str(msg['Subject']) == want['subject'], '제목이 다름: %r' % str(msg['Subject']))
    for k in ('to', 'cc'):
        got = [str(a) for a in (msg[k].addresses if msg[k] else [])]
        exp = [str(Address(display_name=p['name'], addr_spec=p['email'])) for p in want[k]]
        ok(got == exp, '%s 다름: %r != %r' % (k, got, exp))
    ok(all(ord(c) < 128 for c in raw.decode('latin-1')), '메일 글에 ASCII 밖 바이트가 있음(머리글·본문 모두 인코딩해야 함)')
    ok(all(len(l) <= 998 for l in raw.split(b'\r\n')), '998자 넘는 줄')
    ok(b'\n' not in raw.replace(b'\r\n', b''), 'CRLF 아닌 줄바꿈')
    body = msg.get_body(preferencelist=('plain',))
    text = body.get_content() if body else ''
    for s in want['body_has']: ok(s in text, '본문에 없음: ' + s)
    atts = list(msg.iter_attachments())
    ok(len(atts) == 1, '첨부 수 %d' % len(atts))
    if atts:
        a = atts[0]
        ok(a.get_filename() == want['attach'], '첨부 이름: %r' % a.get_filename())
        ok(a.get_content_type() == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', '첨부 형식: ' + a.get_content_type())
        data = a.get_payload(decode=True)
        orig = open(xlsx_path, 'rb').read()
        ok(data == orig, '첨부가 원본과 다름 (%d / %d 바이트)' % (len(data), len(orig)))
        try:
            import openpyxl
            wb = openpyxl.load_workbook(io.BytesIO(data))
            ws = wb['MASTER PLAN']
            rules = []
            for rng in ws.conditional_formatting:
                for r in rng.rules:
                    rules.append((str(rng.sqref), r.type, r.operator, r.formula[0], r.dxf.fill.bgColor.rgb if r.dxf and r.dxf.fill else None))
            exp = [('완료', 'FF1F9D55'), ('진행중', 'FFE8830C'), ('미완료', 'FFC0392B'), ('보류', 'FF7F8C8D')]
            ok([(f.strip('"'), c) for (_, t, op, f, c) in rules] == exp, '조건부 서식: %r' % rules)
            ok(all(t == 'cellIs' and op == 'equal' for (_, t, op, f, c) in rules), '조건부 서식 종류')
            ok(len({s for (s, *_) in rules}) == 1 and rules and rules[0][0].startswith('L'), '조건부 서식 범위: %r' % {s for (s, *_) in rules})
            heads = [c.value for c in ws[5]]
            ok('진행상태' in heads and not any('진행률' in str(h) for h in heads), 'MASTER PLAN 머리: %r' % heads)
            dv = [d for d in ws.data_validations.dataValidation if str(d.sqref) == rules[0][0]] if rules else []
            ok(dv and dv[0].formula1 == '"진행중,완료,미완료,보류"', '진행상태 고르기 목록: %r' % [d.formula1 for d in dv])
        except ImportError:
            print('  (openpyxl 없음 — 엑셀 조건부 서식 검사는 건너뜀)')
        except Exception as e:
            errs.append('첨부 엑셀을 openpyxl 로 열지 못함: %r' % e)
    if errs:
        print('\n'.join('  FAIL ' + e for e in errs)); sys.exit(1)
    print('  ok  .eml — 머리글(RFC 2047) · X-Unsent · 받는 사람/참조 · 본문 · 첨부 %d바이트 원본과 같음 · 조건부 서식 4개' % len(atts[0].get_payload(decode=True)))

if __name__ == '__main__':
    main(*sys.argv[1:4])
