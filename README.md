# eastend-invite

이스트엔드 행사 초대장(모바일) 정적 페이지. 회신은 Google Apps Script 웹앱을 거쳐 **비공개 구글 시트**에만 저장된다.
이 레포에는 명단·회신 같은 개인정보를 두지 않는다.

| 경로 | 내용 |
|---|---|
| `docs/private-edit/index.html` | THE PRIVATE EDIT(2026-11-25) 초대장. `?g=<코드>` 개인 링크면 이름이 미리 찍힘 |
| `docs/private-edit/dashboard.html` | 내부 초청 현황(관리자 토큰 필요) |
| `docs/private-edit/config.js` | 웹앱 `/exec` 주소, 초대장 주소 |
| `gas/Code.gs` | 회신 API 소스(실제 반영은 Apps Script 편집기에서 새 버전 배포) |
