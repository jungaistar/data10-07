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
    "landing": {
      "company": {
        "name": "HD현대",
        "en": "HD HYUNDAI"
      },
      "colors": {
        "navy": "#002554",
        "navy2": "#001a3d",
        "brand": "#003087",
        "brandDark": "#002554",
        "accent": "#5fe35f"
      }
    }
  }
};
