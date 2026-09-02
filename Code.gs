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
    const formType = data.formType || 'training';

    if (formType === 'combined_training') {
      return handleCombinedTraining_(data);
    }

    if (formType === 'class_feedback') {
      return handleClassFeedback_(data);
    }

    const sheetName = sanitizeSheetName_(data.department || data.trainingTitle || DEFAULT_SHEET_NAME);
    const sheet = getOrCreateSheet_(sheetName, formType);
    const folder = getOrCreateFolder_();

    // 서명 이미지를 드라이브에 저장
    const now = new Date();
    const stamp = Utilities.formatDate(now, 'Asia/Seoul', 'yyyyMMdd_HHmmss');
    const nameForFile = formType === 'parent_checkin' ? (data.parentName || '무명') : (data.name || '무명');
    const safeName = String(nameForFile).replace(/[^\w가-힣]/g, '');
    const base64 = String(data.signature || '').split(',')[1] || '';
    let signatureUrl = '';

    if (base64) {
      const blob = Utilities.newBlob(Utilities.base64Decode(base64), 'image/png', `${sheetName}_${safeName}_${stamp}.png`);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      signatureUrl = file.getUrl();
    }

    if (formType === 'parent_checkin') {
      // 학부모 출석체크: 학년/반/성명, 학부모 성명, 연락처, 서명, 비고
      sheet.appendRow([
        now,
        data.grade || '',
        data.classNum || '',
        data.studentName || '',
        data.parentName || '',
        data.phone || '',
        signatureUrl,
        data.note || '',
        data.submittedAt || ''
      ]);
    } else {
      // 기존 연수확인 출석체크 (부서별 연수)
      sheet.appendRow([
        now,
        data.department || '',
        data.trainingTitle || '',
        data.note || '',
        data.name || '',
        data.confirmed ? '확인함' : '미확인',
        signatureUrl,
        data.submittedAt || ''
      ]);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success', sheet: sheetName }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// 학부모 수업 공개 참관록: 서명 없이 평점(1~4) + 서술형 소감을 한 시트에 기록
function handleClassFeedback_(data) {
  const sheetName = sanitizeSheetName_(data.department || '학부모수업참관록');
  const sheet = getOrCreateSheet_(sheetName, 'class_feedback');
  const now = new Date();

  sheet.appendRow([
    now,
    data.grade || '',
    data.classNum || '',
    data.studentName || '',
    data.subject || '',
    data.rating1 || '',
    data.rating2 || '',
    data.rating3 || '',
    data.rating4 || '',
    data.impressive || '',
    data.discovery || '',
    data.opinion || '',
    data.oneline || '',
    data.submittedAt || ''
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'success', sheet: sheetName }))
    .setMimeType(ContentService.MimeType.JSON);
}


// departments 배열에 담긴 만큼 각 부서 시트에 나눠서 한 줄씩 기록한다.
function handleCombinedTraining_(data) {
  const folder = getOrCreateFolder_();
  const now = new Date();
  const stamp = Utilities.formatDate(now, 'Asia/Seoul', 'yyyyMMdd_HHmmss');
  const safeName = String(data.parentName || '무명').replace(/[^\w가-힣]/g, '');

  // 서명은 학부모당 1번만 드라이브에 저장하고, 모든 부서 시트에서 같은 링크를 공유한다.
  const base64 = String(data.signature || '').split(',')[1] || '';
  let signatureUrl = '';
  if (base64) {
    const blob = Utilities.newBlob(Utilities.base64Decode(base64), 'image/png', `통합_${safeName}_${stamp}.png`);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    signatureUrl = file.getUrl();
  }

  const departments = Array.isArray(data.departments) ? data.departments : [];
  const savedSheets = [];

  departments.forEach(dept => {
    const sheetName = sanitizeSheetName_(dept.name || dept.code || DEFAULT_SHEET_NAME);
    const sheet = getOrCreateSheet_(sheetName, 'combined_training');
    sheet.appendRow([
      now,
      data.grade || '',
      data.classNum || '',
      data.studentName || '',
      data.parentName || '',
      data.phone || '',
      dept.trainingTitle || '',
      '확인함',
      signatureUrl,
      data.submittedAt || ''
    ]);
    savedSheets.push(sheetName);
  });

  // 부서별 연수 확인과 별개로, 같은 서명/정보로 학부모 출석 기록도 함께 남긴다.
  const attendanceSheet = getOrCreateSheet_('학부모출석체크', 'parent_checkin');
  attendanceSheet.appendRow([
    now,
    data.grade || '',
    data.classNum || '',
    data.studentName || '',
    data.parentName || '',
    data.phone || '',
    signatureUrl,
    data.note || '',
    data.submittedAt || ''
  ]);
  savedSheets.push('학부모출석체크');

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'success', sheets: savedSheets }))
    .setMimeType(ContentService.MimeType.JSON);
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

function getOrCreateSheet_(sheetName, formType) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    if (formType === 'parent_checkin') {
      sheet.appendRow(['제출시각', '학년', '반', '자녀 성명', '학부모 성명', '연락처', '서명이미지', '비고', '클라이언트 제출시각']);
    } else if (formType === 'combined_training') {
      sheet.appendRow(['제출시각', '학년', '반', '자녀 성명', '학부모 성명', '연락처', '연수명', '확인여부', '서명이미지', '클라이언트 제출시각']);
    } else if (formType === 'class_feedback') {
      sheet.appendRow(['제출시각', '학년', '반', '학생명', '참관교과', '평가1_참여', '평가2_표현기회', '평가3_소통협력', '평가4_성장확인', '인상깊었던점', '새롭게발견한점', '학교에전할의견', '한마디표현', '클라이언트 제출시각']);
    } else {
      sheet.appendRow(['제출시각', '부서', '연수명', '자녀/학생', '이름', '확인여부', '서명이미지', '클라이언트 제출시각']);
    }
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
