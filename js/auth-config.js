// 로그인 설정 — 이 파일 한 곳만 바꾸면 된다(생성: data10-00/scripts/apply-auth-kit.mjs, 대문 문구 출처 docs/plan.json).
// 내 Supabase 로 옮길 때: url · anonKey 를 내 프로젝트 값으로, 카카오를 설정하지 않았으면 kakao: false. 안내 = docs/내_깃허브로_옮기기.md
// prefix 는 supabase/auth.sql 의 표 이름 앞부분과 같아야 한다.
window.AUTH_CONFIG = {
  "url": "https://hcmgdztsgjvzcyxyayaj.supabase.co",
  "anonKey": "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6ImhjbWdkenRzZ2p2emN5eHlheWFqIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzE0MzU4ODcsImV4cCI6MjA4NzAxMTg4N30.gznaPzY1l8qDAPsEyYNR9KS7f7VqS3xaw-_2HTSwSZw",
  "prefix": "data1007",
  "kakao": true,
  "idleMinutes": 30,
  "localBypass": true,
  "board": true,
  "site": {
    "title": "협력사 공정불량 Master Plan",
    "summary": "협력사 공정불량을 등록하거나 캠퍼스별 RAW DATA 를 올려 불량유형 10종으로 나누고, WORST 1·2·3 · 동일품번 반복을 우선 개선 항목으로 개선 MASTER PLAN 초안과 간트를 만들어 협력사와 Excel · 메일로 진행사항을 주고받는 도구입니다.",
    "course": "현장 데이터 디지털화 전문가과정 · 2차수",
    "features": [
      {
        "name": "불량 등록 · 캠퍼스 RAW",
        "desc": "한 건씩 적거나 엑셀 표를 붙여 넣고, 울산 · 인천캠퍼스 RAW DATA 를 각각 올려 합칩니다."
      },
      {
        "name": "불량 분석",
        "desc": "불량유형 10종으로 나눠 파레토 · 월별 추이 · 교차표를 보여 주고 WORST 1·2·3 · 동일품번을 뽑습니다."
      },
      {
        "name": "MASTER PLAN · 간트",
        "desc": "유형별 기본 대책으로 초안을 만들고 간트에서 계획 · 실적과 진행상태 4가지를 색으로 관리합니다."
      },
      {
        "name": "Excel · 메일 · 회신",
        "desc": "두 시트 Excel 과 아웃룩 메일 초안(.eml)으로 보내고, 협력사 회신 Excel 의 바뀐 칸만 반영합니다."
      },
      {
        "name": "추진현황",
        "desc": "모든 협력사의 개선과제를 진행상태별로 거르고 협력사별 표와 간트로 봅니다."
      },
      {
        "name": "Copilot 로 다듬기",
        "desc": "협력사 · 품번을 별칭으로 가린 프롬프트를 복사해 Copilot 등에 붙여 넣고 답을 받아 반영합니다."
      }
    ],
    "steps": [
      "로그인 · 승인",
      "불량 넣기",
      "유형 분류",
      "개선 항목",
      "MASTER PLAN",
      "Excel · 메일"
    ],
    "flow": [
      {
        "step": "로그인 · 승인",
        "desc": "메일로 가입 → 사용 신청 → 관리자 승인(사내 PC 의 HTML 한 파일은 로그인 없음)"
      },
      {
        "step": "불량 넣기",
        "desc": "화면 등록 · 엑셀 붙여넣기 · 울산 · 인천캠퍼스 RAW 올리기"
      },
      {
        "step": "유형 분류",
        "desc": "직접 고른 유형 → 붙여 넣은 유형 칸 → 내용 키워드 순으로 10종"
      },
      {
        "step": "개선 항목",
        "desc": "협력사별 WORST 1·2·3 · 동일품번 2건 이상"
      },
      {
        "step": "MASTER PLAN",
        "desc": "기본 대책 초안 → Copilot 로 다듬기 → 간트로 계획 확정"
      },
      {
        "step": "Excel · 메일",
        "desc": "두 시트 Excel, 그 엑셀을 붙인 .eml 메일 초안"
      }
    ],
    "goals": [
      "공정불량을 불량유형 10종으로 자동 분류하고 협력사별 발생 현황을 보여 줍니다.",
      "불량유형 WORST 1·2·3 과 동일품번 2건 이상을 우선 개선 항목으로 뽑아 MASTER PLAN 초안을 만듭니다.",
      "간트로 계획 · 실적과 진행상태 4가지(진행중 · 완료 · 미완료 · 보류)를 색으로 관리합니다."
    ],
    "screen": "docs/img/screen.png",
    "plan": "plan.html",
    "home": true,
    "guide": {
      "screens": [
        {
          "name": "대시보드",
          "desc": "불량 발생 건수·협력사 수·개선 필요 항목·개선과제, 개선 MASTER PLAN 현황(전체·진행중·완료·미완료·보류), 협력사별 건수, 불량유형 파레토, 월별 추이, 개선 필요 항목 TOP 8(구분 WORST·동일품번), 협력사별 추진현황, 일하는 순서"
        },
        {
          "name": "불량 등록",
          "desc": "기본 경로. ① 한 건씩 등록(발생일자·협력사·품번·품명·불량유형 8종·수량·내용·조치, 유형을 비우면 내용 키워드로 자동 분류 — 입력 중 미리 보임) ② 엑셀에서 붙여넣기(머리행이 있으면 열 이름으로, 없으면 「발생일자·협력사·품번·품명·불량유형·내용·수량·조치」 순서) ③ 등록된 불량 — 거르기·찾기·수정·지우기 ④ 백업·복원(JSON·엑셀) · 입력 양식(.xlsx)"
        },
        {
          "name": "파일 업로드 (선택)",
          "desc": "울산캠퍼스 · 인천캠퍼스 RAW DATA 칸(그 캠퍼스만 바꾸기) + 그 밖의 Excel·CSV — 제목 줄 아래 머리행 자동 찾기, 다른 열 이름 자동 연결, 직접 고르기"
        },
        {
          "name": "불량 분석",
          "desc": "협력사·캠퍼스·기간 거르기 · 유형 파레토 · 월별 추이 · 협력사×유형 교차표 · 개선 필요 항목(① 불량유형 WORST 1·2·3 표 ② 동일품번 2건 이상 표) · 분류 결과(행마다 유형 직접 고르기, 키워드 바로 더하기)"
        },
        {
          "name": "MASTER PLAN",
          "desc": "협력사·계획 착수 기준일·분석기간 → 「분석 및 초안 생성」(기본 대책 규칙) · 지표 6개 · 불량 분석 요약 · 통합 개선대책 Gantt(주차/월, 계획·실적 막대, 상태 색, 오늘 선) · 행을 누르면 오른쪽 칸에서 계획·실적 입력 · 「개선활동 추가」「계획 확정」「초안 다시 불러오기」「Excel 보고서」「메일 만들기(아웃룩)」 · AI 로 다듬기(선택, 접혀 있음 — 협력사·품번은 별칭으로 가림)"
        },
        {
          "name": "추진현황",
          "desc": "개선 MASTER PLAN 현황 타일, 진행상태(진행중·완료·미완료·보류) 거르기, 협력사별 표, 간트, 행 눌러 실적 입력"
        },
        {
          "name": "보고서 · Excel 추출",
          "desc": "협력사 하나 또는 전체(캠퍼스 거르기) — 「불량목록」(캠퍼스 · 관리번호 칸)·「MASTER PLAN」 두 시트(노란 칸 = 협력사가 고치는 칸, 숨긴 ID·기준 칸, 진행상태 고르기 목록 + 조건부 서식 색), 메일 만들기, 인쇄/PDF"
        },
        {
          "name": "진행사항 업로드",
          "desc": "협력사 회신 Excel → ID 로 행 맞추기 → 바뀐 칸·새 활동·충돌·확인 필요 미리 보기 → 반영 → 이력(일시·파일·협력사·반영한 사람·바뀐 내용)"
        },
        {
          "name": "설정",
          "desc": "HTML 한 파일 받기, 메일 받는 사람 · 초안(협력사별 받는 사람), 불량유형 10종 키워드·기본 대책(순서 = 우선순위), 담당부서(협력사 내부, 기본 5종), 개선 필요 기준(WORST 몇 위 · 동일품번 몇 건), 별칭 가리기, 사내망 안내, AI 연결(선택), 백업·복원"
        }
      ],
      "data": [
        {
          "name": "공정불량",
          "fields": "캠퍼스 · 관리번호 · 발생일자 · 협력사 · 품번 · 품명 · 불량유형 · 내용 · 수량 · 조치"
        },
        {
          "name": "캠퍼스 RAW DATA",
          "fields": "울산(문제점 번호 · 발생 일자 · 협력회사명 · 문제점 상세 …) · 인천(QM noti No. · Issued Date · Vendor Name · Problem Description …)"
        },
        {
          "name": "MASTER PLAN",
          "fields": "협력사 · 불량유형 · 개선대책 · 담당부서 · 계획 시작 · 마감 · 실적 · 진행상태"
        },
        {
          "name": "내보내기",
          "fields": "「불량목록」 · 「MASTER PLAN」 Excel · 메일 .eml · JSON 백업 · HTML 한 파일"
        }
      ],
      "rules": [
        "불량유형은 직접 고른 유형 → 붙여 넣은 유형 칸 → 내용 키워드 순으로 정합니다. 어느 것에도 맞지 않으면 미분류로 남겨 고르게 합니다.",
        "개선 필요 항목 = 협력사별 불량유형 WORST 1·2·3 과 같은 품번 2건 이상(설정에서 바꿈).",
        "진행상태: 완료 · 보류는 적은 그대로, 나머지는 개선 마감이 지났으면 미완료, 아니면 진행중(시작 전도 진행중).",
        "중복은 같은 캠퍼스 안에서만 봅니다. 캠퍼스 RAW 를 다시 올리면 그 캠퍼스에서 올린 행만 바뀝니다.",
        "MASTER PLAN 과 WORST 는 협력사 단위(캠퍼스를 합쳐 셈)입니다.",
        "회신 Excel 은 숨긴 ID 칸으로 행을 맞추고, 같은 파일을 다시 올려도 한 번만 반영합니다."
      ],
      "limits": [
        "업무 자료(불량 · 계획 · 이력)는 아직 팀 공유 DB 가 아니라 이 브라우저에 저장됩니다. 나누려면 JSON · 엑셀 백업을 주고받습니다.",
        "개선대책 초안은 규칙(기본 대책)으로 만들고, AI 는 프롬프트 복사 · 붙여 넣기가 기본입니다.",
        "브라우저는 아웃룩을 직접 열거나 보내지 못해 .eml 파일로 건넵니다.",
        "협력사가 사이트에 직접 들어와 입력하는 기능은 없습니다(Excel 회신으로 받음)."
      ],
      "users": [
        "부품품질팀 담당자",
        "협력사 품질 담당(회신 Excel 작성)"
      ],
      "flow": [
        {
          "step": "로그인 · 승인",
          "desc": "메일로 가입 → 사용 신청 → 관리자 승인(사내 PC 의 HTML 한 파일은 로그인 없음)"
        },
        {
          "step": "불량 넣기",
          "desc": "화면 등록 · 엑셀 붙여넣기 · 울산 · 인천캠퍼스 RAW 올리기"
        },
        {
          "step": "유형 분류",
          "desc": "직접 고른 유형 → 붙여 넣은 유형 칸 → 내용 키워드 순으로 10종"
        },
        {
          "step": "개선 항목",
          "desc": "협력사별 WORST 1·2·3 · 동일품번 2건 이상"
        },
        {
          "step": "MASTER PLAN",
          "desc": "기본 대책 초안 → Copilot 로 다듬기 → 간트로 계획 확정"
        },
        {
          "step": "Excel · 메일",
          "desc": "두 시트 Excel, 그 엑셀을 붙인 .eml 메일 초안"
        },
        {
          "step": "회신 반영",
          "desc": "협력사 회신 Excel 의 바뀐 칸 미리 보기 → 반영 → 이력"
        }
      ]
    },
    "landing": {
      "company": {
        "name": "HD건설기계",
        "en": "HD CONSTRUCTION EQUIPMENT",
        "logo": "images/logo.png",
        "logoH": 24
      },
      "colors": {
        "navy": "#002554",
        "navy2": "#001a3d",
        "brand": "#003087",
        "brandDark": "#002554",
        "accent": "#5fe35f"
      },
      "slides": [
        {
          "image": "images/slide-1.jpg"
        },
        {
          "image": "images/slide-2.jpg"
        },
        {
          "image": "images/slide-3.jpg"
        }
      ],
      "heroMode": "slide"
    }
  }
};
