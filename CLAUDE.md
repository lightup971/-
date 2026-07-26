# 저장소 가이드

## 채용 지원 자동화 (job-hunt/)

개인 취업 비서 시스템. 스킬 3개로 채용 지원 전 과정을 자동화합니다.
전체 설명은 `job-hunt/README.md` 참고.

- **"브리핑"** → `daily-briefing` 스킬 (공고 매칭 + 마감 관리)
- **"OO 자소서 써줘"** → `cover-letter` 스킬 (6단계 초안 생성)
- **"OO 제출했어"** → `submit-done` 스킬 (트래커/캘린더 동시 업데이트)

기준 데이터: `job-hunt/profile.md`(매칭 기준), `job-hunt/experience-blocks.md`(자소서 재료),
`job-hunt/tracker.md`(지원 현황). 스킬 실행 전 이 파일들을 먼저 읽습니다.

## 부릉 사업소득 자동화 (bureung-automation/)
별개 도구. 해당 폴더 README 참고.
