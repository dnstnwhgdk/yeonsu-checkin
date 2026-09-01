/**
 * 연수확인 출석체크 - 구글 앱스 스크립트 백엔드
 * (부서별로 시트 탭을 자동 분리해서 저장)
 *
 * [사용 방법]
 * 1. 구글 워크시트(스프레드시트)를 하나 새로 만든다.
 * 2. 확장 프로그램 > Apps Script > 이 코드 전체 붙여넣기.
 * 3. DRIVE_FOLDER_ID는 비워둬도 동작함 (서명 이미지 저장용 드라이브 폴더).
 * 4. 배포 > 새 배포 > 유형 "웹 앱" 선택
 *    - 실행 계정: 나
 *    - 액세스 권한이 있는 사용자: 모든 사용자
 * 5. 배포 후 발급되는 웹 앱 URL을 각 부서 index.html의 SCRIPT_URL에 붙여넣는다.
 *
 * 각 부서 페이지(index.html)는 DEPARTMENT 상수를 함께 전송하며,
 * 이 값과 동일한 이름의 시트 탭이 자동으로 만들어져 그 안에 기록된다.
 * (탭이 없으면 새로 생성, 있으면 이어서 추가)
 */

const DRIVE_FOLDER_ID = ''; // 서명 이미지를 저장할 구글 드라이브 폴더 ID (비워두면 자동 생성)
const DEFAULT_SHEET_NAME = '기타'; // department 값이 없을 때 사용할 시트 이름

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);

    const sheetName = sanitizeSheetName_(data.department || data.trainingTitle || DEFAULT_SHEET_NAME);
    const sheet = getOrCreateSheet_(sheetName);
    const folder = getOrCreateFolder_();

    // 서명 이미지를 드라이브에 저장
    const now = new Date();
    const stamp = Utilities.formatDate(now, 'Asia/Seoul', 'yyyyMMdd_HHmmss');
    const safeName = String(data.name || '무명').replace(/[^\w가-힣]/g, '');
    const base64 = String(data.signature || '').split(',')[1] || '';
    let signatureUrl = '';

    if (base64) {
      const blob = Utilities.newBlob(Utilities.base64Decode(base64), 'image/png', `${sheetName}_${safeName}_${stamp}.png`);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      signatureUrl = file.getUrl();
    }

    sheet.appendRow([
      now,
      data.department || '',
      data.trainingTitle || '',
      data.name || '',
      data.confirmed ? '확인함' : '미확인',
      signatureUrl,
      data.submittedAt || ''
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success', sheet: sheetName }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok', message: '연수확인 출석체크 서버가 정상 동작 중입니다.' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// 시트(스프레드시트) 탭 이름 제약 처리: 100자 제한, 대괄호/물음표/슬래시 등 금지문자 제거
function sanitizeSheetName_(name) {
  let clean = String(name)
    .replace(/[\[\]\*\?\/\\:]/g, '') // 금지 문자 제거
    .trim();
  if (clean.length > 90) clean = clean.substring(0, 90); // 여유 있게 90자로 컷
  return clean || DEFAULT_SHEET_NAME;
}

function getOrCreateSheet_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['제출시각', '부서', '연수명', '이름', '확인여부', '서명이미지', '클라이언트 제출시각']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateFolder_() {
  if (DRIVE_FOLDER_ID) {
    return DriveApp.getFolderById(DRIVE_FOLDER_ID);
  }
  const folderName = '연수확인_서명이미지';
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(folderName);
}
