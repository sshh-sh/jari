/**
 * 🎲 모둠뽑기 - Google Apps Script 백엔드
 */

const SHEET_NAME   = '모둠뽑기';      // 앱 내부용 (JSON 원본)
const SHEET_DETAIL = '모둠뽑기_보기'; // 결과 + 모둠장 횟수 통합 시트

// ===== 웹앱 진입점 =====
function doPost(e) {
  try {
    const data    = JSON.parse(e.postData.contents);
    const action  = data.action;
    const classId = data.classId;
    const payload = data.data;

    let result;
    if (action === 'save') result = saveRecord(classId, payload);
    else if (action === 'load') result = loadRecords(classId);
    else result = {success: false, message: '알 수 없는 액션'};

    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({success: false, message: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({status:'ok', message:'모둠뽑기 GAS 작동 중'})).setMimeType(ContentService.MimeType.JSON);
}

// ===== 기록 저장 =====
function saveRecord(classId, record) {
  const sheet = getOrCreateSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === classId && data[i][1] === record.year && data[i][2] === record.month) {
      sheet.getRange(i+1,1,1,5).setValues([[classId, record.year, record.month, record.date, JSON.stringify(record.groups)]]);
      saveDetailSheet(classId, record);
      return {success: true, message: '기존 기록 업데이트 완료'};
    }
  }

  sheet.appendRow([classId, record.year, record.month, record.date, JSON.stringify(record.groups)]);
  saveDetailSheet(classId, record);
  return {success: true, message: '저장 완료'};
}

// ===== 통합 시트 저장 =====
function saveDetailSheet(classId, record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_DETAIL);

  const maxGroups = 6;
  // A~I: 결과, J: 빈칸(구분선), K: 번호, L: 이름, M: 모둠장 횟수
  const resultHeaders = ['반', '년도', '월', '1모둠', '2모둠', '3모둠', '4모둠', '5모둠', '6모둠'];
  const leaderHeaders = ['', '번호', '이름', '모둠장 횟수'];

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_DETAIL);
    // 결과 헤더 (A1~I1)
    sheet.getRange(1,1,1,9).setValues([resultHeaders]);
    sheet.getRange(1,1,1,9).setFontWeight('bold').setBackground('#534AB7').setFontColor('white');
    sheet.setFrozenRows(1);
    for (let i = 4; i <= 9; i++) sheet.setColumnWidth(i, 180);
  }

  const className = record.className || classId;

  // 같은 반+월 기존 행 삭제
  const existing = sheet.getDataRange().getValues();
  const toDelete = [];
  for (let i = existing.length-1; i >= 1; i--) {
    if (existing[i][0] === className && existing[i][1] === record.year && existing[i][2] === record.month) {
      toDelete.push(i+1);
    }
  }
  toDelete.forEach(row => sheet.deleteRow(row));

  // 결과 행 추가 (A~I)
  const row = [className, record.year, record.month];
  for (let i = 0; i < maxGroups; i++) {
    const group = record.groups[i];
    row.push(group ? group.students.join(', ') : '');
  }
  sheet.appendRow(row);

  // 배경색
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow,1,1,9).setBackground(lastRow%2===0 ? '#f0effe' : '#ffffff');

  // 모둠장 횟수 업데이트 (L열에 이름 있으면 M열 자동 갱신)
  if (record.leaders && record.leaders.length > 0) {
    updateLeaderCount(sheet, record.leaders, className);
  }
}

// ===== 모둠장 횟수 업데이트 =====
// J1, N1, R1, V1... 에 반 이름 써두면 자동으로 찾아서 해당 반 횟수 업데이트
// 패턴: 반이름(col), 이름(col+1), 횟수(col+2), 구분(col+3) → 4칸 간격
function updateLeaderCount(sheet, leaders, className) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // row 1 전체 읽어서 반 이름 위치 찾기 (J=10번째 컬럼부터 4칸 간격)
  const lastCol = sheet.getLastColumn();
  const headerRow = sheet.getRange(1, 1, 1, Math.max(lastCol, 10)).getValues()[0];

  let classCol = -1;
  for (let c = 9; c < headerRow.length; c += 4) { // J열=index 9, 4칸 간격
    if (String(headerRow[c]).trim() === String(className).trim()) {
      classCol = c + 1; // 1-indexed
      break;
    }
  }
  if (classCol === -1) return; // 해당 반 없으면 종료

  const nameCol  = classCol + 1; // 이름 열
  const countCol = classCol + 2; // 횟수 열

  // 이름열, 횟수열 읽기
  const nameValues  = sheet.getRange(2, nameCol,  lastRow-1, 1).getValues();
  const countValues = sheet.getRange(2, countCol, lastRow-1, 1).getValues();

  leaders.forEach(leaderName => {
    for (let i = 0; i < nameValues.length; i++) {
      if (String(nameValues[i][0]).trim() === String(leaderName).trim()) {
        countValues[i][0] = (Number(countValues[i][0]) || 0) + 1;
      }
    }
  });

  sheet.getRange(2, countCol, lastRow-1, 1).setValues(countValues);
}

// ===== 기록 불러오기 =====
function loadRecords(classId) {
  const sheet = getOrCreateSheet();
  const data  = sheet.getDataRange().getValues();
  const records = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === classId) {
      try {
        records.push({year:data[i][1], month:data[i][2], date:data[i][3], groups:JSON.parse(data[i][4])});
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
    sheet.getRange(1,1,1,5).setValues([['반 ID','년도','월','날짜','모둠 데이터(JSON)']]);
    sheet.setFrozenRows(1);
    sheet.getRange(1,1,1,5).setFontWeight('bold').setBackground('#534AB7').setFontColor('white');
    sheet.setColumnWidth(5,400);
  }

  return sheet;
}
