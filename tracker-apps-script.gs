/**
 * 증상 기록 트래커 — 구글시트 저장용 Apps Script
 *
 * 사용법:
 * 1. sheets.google.com 에서 새 스프레드시트를 만든다 (감사일기용과는 다른, 별도의 시트를 추천)
 * 2. 상단 메뉴 [확장 프로그램] → [Apps Script] 클릭
 * 3. 기본으로 있던 코드를 지우고 이 파일 내용 전체를 붙여넣는다
 * 4. 저장(Ctrl+S) 후 [배포] → [새 배포] 클릭
 * 5. 유형: "웹 앱" 선택
 *    - 실행할 계정: 나
 *    - 액세스 권한이 있는 사용자: 전체
 * 6. [배포] 클릭 → 나오는 웹 앱 URL을 복사해서 앱의 "구글 Apps Script 웹앱 URL" 칸에 붙여넣는다
 * 7. 처음 배포할 때 구글이 권한 승인을 요청하면 "고급" → "이동(안전하지 않음)" 눌러서 허용해준다
 *    (내가 만든 스크립트라서 안전합니다)
 */

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = JSON.parse(e.postData.contents);

  // 첫 실행이라 헤더가 없으면 추가
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['날짜', '시간', '증상', '상세', '등급', '저장된 시각']);
  }

  sheet.appendRow([
    data.dateStr || '',
    data.time || '',
    data.label || '',
    data.detail || '',
    data.grade || '',
    Utilities.formatDate(new Date(), 'GMT+9', 'yyyy-MM-dd HH:mm:ss')
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService.createTextOutput('증상 기록 저장 서버가 정상 작동 중입니다.');
}
