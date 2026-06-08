/**
 * 🎲 모둠뽑기 - Google Apps Script 백엔드
 */

const SHEET_NAME   = '모둠기록';      // 앱 내부용 (JSON 원본)
const SHEET_DETAIL = '모둠기록_보기'; // 보기 좋은 형태

// ===== 웹앱 진입점 =====
function doPost(e) {
  try {
    const data    = JSON.parse(e.postData.contents);
    const action  = data.action;
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
      saveDetailSheet(classId, record);
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

  saveDetailSheet(classId, record);
  return {success: true, message: '저장 완료'};
}

// ===== 보기 좋은 시트에 저장 =====
function saveDetailSheet(classId, record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_DETAIL);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_DETAIL);
    const headers = ['반', '년도', '월', '모둠명', '학생 명단'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold')
      .setBackground('#534AB7')
      .setFontColor('white');
    sheet.setColumnWidth(5, 350);
  }

  // 반 이름 (className 없으면 classId 그대로)
  const className = record.className || classId;

  // 같은 반+월 기존 행 삭제
  const existing = sheet.getDataRange().getValues();
  const toDelete = [];
  for (let i = existing.length - 1; i >= 1; i--) {
    if (existing[i][0] === className && existing[i][1] === record.year && existing[i][2] === record.month) {
      toDelete.push(i + 1);
    }
  }
  toDelete.forEach(row => sheet.deleteRow(row));

  // 모둠별로 한 행씩 추가
  record.groups.forEach((group, i) => {
    sheet.appendRow([
      className,
      record.year,
      record.month,
      group.name,
      group.students.join(', ')
    ]);
  });

  // 모둠별 번갈아 배경색 (가독성)
  const lastRow = sheet.getLastRow();
  const startRow = lastRow - record.groups.length + 1;
  record.groups.forEach((group, i) => {
    const color = i % 2 === 0 ? '#f0effe' : '#ffffff';
    sheet.getRange(startRow + i, 1, 1, 5).setBackground(color);
  });
}

// ===== 기록 불러오기 =====
function loadRecords(classId) {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
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
    sheet.getRange(1, 1, 1, 5)
      .setFontWeight('bold')
      .setBackground('#534AB7')
      .setFontColor('white');
    sheet.setColumnWidth(5, 400);
  }

  return sheet;
}
