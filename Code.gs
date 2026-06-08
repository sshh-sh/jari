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
    // 모둠장 헤더 (J1~M1)
    sheet.getRange(1,10,1,4).setValues([leaderHeaders]);
    sheet.getRange(1,11,1,3).setFontWeight('bold').setBackground('#1e3a30').setFontColor('white');
    sheet.setFrozenRows(1);
    // 컬럼 너비
    for (let i = 4; i <= 9; i++) sheet.setColumnWidth(i, 180);
    sheet.setColumnWidth(10, 20);  // J: 구분선
    sheet.setColumnWidth(11, 60);  // K: 번호
    sheet.setColumnWidth(12, 100); // L: 이름
    sheet.setColumnWidth(13, 100); // M: 모둠장 횟수
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
    updateLeaderCount(sheet, record.leaders);
  }
}

// ===== 모둠장 횟수 업데이트 (L열 이름 기준으로 M열 갱신) =====
function updateLeaderCount(sheet, leaders) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // L열(12번째) 전체 읽기
  const nameRange = sheet.getRange(2, 12, lastRow-1, 1).getValues(); // L열: 이름
  const countRange = sheet.getRange(2, 13, lastRow-1, 1).getValues(); // M열: 횟수

  leaders.forEach(leaderName => {
    for (let i = 0; i < nameRange.length; i++) {
      if (nameRange[i][0] === leaderName) {
        countRange[i][0] = (countRange[i][0] || 0) + 1;
      }
    }
  });

  // M열 업데이트
  sheet.getRange(2, 13, lastRow-1, 1).setValues(countRange);
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
