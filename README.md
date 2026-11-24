# 위더스 일정비서

에어컨 청소·집 청소 예약 문자를 받으면 일정표와 비교해서, 겹치면 알려주고 대신 가능한 시간과 답장 문자를 만들어 주는 시스템이에요.

| 폴더 | 내용 |
|---|---|
| `apps-script/` | **실제로 쓰는 버전.** 구글 시트 + Apps Script로 돌아가요 (무료). 설치는 [`apps-script/SETUP.md`](apps-script/SETUP.md) |
| `pwa/ai-schedule.html` | 체험용 데모 페이지 (예시 데이터, 서버 없음) |
| `server.js` 등 | 예전 구글 설문지 + Render 서버 버전 |

예전 Render 무료 서버 버전은 서버가 재시작될 때마다 SQLite 파일이 지워지고, 한동안 안 쓰면 잠들어서 첫 응답이 느려요. 그래서 실제 버전은 구글 Apps Script로 옮겼어요.
