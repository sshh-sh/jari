/**
 * 🎲 모둠뽑기 - Google Apps Script 백엔드
 */

const SHEET_NAME   = '모둠뽑기';      // 앱 내부용 (JSON 원본)
const SHEET_DETAIL = '모둠뽑기_보기'; // 보기 좋은 형태
const SHEET_LEADER = '모둠장 횟수';   // 모둠장 누적 카운트

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
      if (record.leaders && record.leaders.length > 0) updateLeaderSheet(record.className || classId, record.leaders);
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
  if (record.leaders && record.leaders.length > 0) updateLeaderSheet(record.className || classId, record.leaders);
  return {success: true, message: '저장 완료'};
}

// ===== 보기 좋은 시트에 저장 =====
function saveDetailSheet(classId, record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_DETAIL);

  const maxGroups = 6;
  const headers = ['반', '년도', '월'];
  for (let i = 1; i <= maxGroups; i++) headers.push(i + '모둠');
  const totalCols = headers.length;

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_DETAIL);
    sheet.getRange(1, 1, 1, totalCols).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, totalCols)
      .setFontWeight('bold')
      .setBackground('#534AB7')
      .setFontColor('white');
    for (let i = 4; i <= totalCols; i++) sheet.setColumnWidth(i, 200);
  }

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

  // 한 행에 모든 모둠 저장
  const row = [className, record.year, record.month];
  for (let i = 0; i < maxGroups; i++) {
    const group = record.groups[i];
    row.push(group ? group.students.join(', ') : '');
  }
  sheet.appendRow(row);

  // 배경색
  const lastRow = sheet.getLastRow();
  const color = lastRow % 2 === 0 ? '#f0effe' : '#ffffff';
  sheet.getRange(lastRow, 1, 1, totalCols).setBackground(color);
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

// ===== 모둠장 횟수 업데이트 =====
function updateLeaderSheet(className, leaders) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_LEADER);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_LEADER);
    sheet.getRange(1,1,1,3).setValues([['반', '이름', '모둠장 횟수']]);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,3).setFontWeight('bold').setBackground('#534AB7').setFontColor('white');
    sheet.setColumnWidth(2, 120);
    sheet.setColumnWidth(3, 120);
  }

  const data = sheet.getDataRange().getValues();

  leaders.forEach(name => {
    // 같은 반 + 같은 이름 찾기
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === className && data[i][1] === name) {
        const newCount = (data[i][2] || 0) + 1;
        sheet.getRange(i+1, 3).setValue(newCount);
        data[i][2] = newCount; // 로컬 업데이트
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([className, name, 1]);
      data.push([className, name, 1]);
    }
  });

  // 횟수 내림차순 정렬
  const lastRow = sheet.getLastRow();
  if (lastRow > 2) {
    const range = sheet.getRange(2, 1, lastRow-1, 3);
    range.sort({column: 3, ascending: false});
  }
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
