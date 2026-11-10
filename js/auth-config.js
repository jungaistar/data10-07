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
      }
    ],
    "steps": [
      "로그인 · 승인",
      "불량 넣기",
      "유형 분류",
      "개선 항목",
      "MASTER PLAN",
      "Excel · 메일"
    ]
  }
};
