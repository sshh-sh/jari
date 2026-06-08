/**
 * 🎲 모둠뽑기 - Google Apps Script 백엔드
 * 스프레드시트에 배치 결과를 저장/불러오는 역할을 합니다.
 *
 * 사용 방법:
 * 1. Google 스프레드시트 새로 만들기
 * 2. 확장 프로그램 → Apps Script
 * 3. 이 코드 전체를 붙여넣기
 * 4. 저장 → 배포 → 새 배포 → 웹앱으로 배포
 * 5. 액세스 권한: '모든 사용자' 설정
 * 6. 배포 URL을 복사해서 앱의 'Google 저장소 연결'에 붙여넣기
 */

const SHEET_NAME = '모둠기록';

// ===== 웹앱 진입점 =====
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const classId = data.classId;
    const payload = data.data;

    let result;
    if (action === 'save') {
      result = saveRecord(classId, payload);
    } else if (action === 'load') {
      result = loadRecords(classId);
    } else {
      result = {success: false, message: '알 수 없는 액션'};
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({success: false, message: err.toString()}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({status: 'ok', message: '모둠뽑기 GAS 작동 중'}))
    .setMimeType(ContentService.MimeType.JSON);
}

// ===== 기록 저장 =====
function saveRecord(classId, record) {
  const sheet = getOrCreateSheet();

  // 같은 반 + 같은 월 기록이 있으면 덮어쓰기
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === classId && data[i][1] === record.year && data[i][2] === record.month) {
      sheet.getRange(i + 1, 1, 1, 5).setValues([[
        classId,
        record.year,
        record.month,
        record.date,
        JSON.stringify(record.groups)
      ]]);
      return {success: true, message: '기존 기록 업데이트 완료'};
    }
  }

  // 새 행 추가
  sheet.appendRow([
    classId,
    record.year,
    record.month,
    record.date,
    JSON.stringify(record.groups)
  ]);

  return {success: true, message: '저장 완료'};
}

// ===== 기록 불러오기 =====
function loadRecords(classId) {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();
  const records = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === classId) {
      try {
        records.push({
          year:   data[i][1],
          month:  data[i][2],
          date:   data[i][3],
          groups: JSON.parse(data[i][4])
        });
      } catch(e) {}
    }
  }

  return {success: true, records};
}

// ===== 시트 가져오거나 생성 =====
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, 5).setValues([['반 ID', '년도', '월', '날짜', '모둠 데이터(JSON)']]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#534AB7').setFontColor('white');
    sheet.setColumnWidth(5, 400);
  }

  return sheet;
}
