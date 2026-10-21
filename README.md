# a4 — 지점별 A4 발주 취합 (화면)

GitHub Pages 로 서비스되는 정적 화면입니다. 데이터는 저장소에 없고, Google Apps Script 웹앱(`js/api.js`의 `A4.API_URL`)을 통해 스프레드시트에 저장됩니다.

- `index.html` — 주임단 확인·수정 화면 (`?k=접속코드`)
- `admin.html` — 담당자 관리 화면 (파일 업로드·집계·회차 생성·현황·엑셀 다운로드)
- `js/agg.js` — 집계 모듈 (브라우저/Node 공용)
- `js/api.js` — API 클라이언트·공통 유틸
- `css/app.css` — 스타일

원본 소스와 기획서는 `project/4.rental` 에 있습니다.
