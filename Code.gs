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
// 시트 구조: J(반+번호) K(횟수) L(이름) | N(반+번호) O(횟수) P(이름) | R S T | V W X
// "3학년 1반" → "3-1" 로 변환해서 J열에서 클래스 마커 찾기
function updateLeaderCount(sheet, leaders, className) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // "3학년 1반" → grade=3, classNum=1 → marker="3-1"
  const gradeMatch = className.match(/(\d+)학년/);
  const classNumMatch = className.match(/(\d+)반/);
  if (!gradeMatch || !classNumMatch) return;
  const grade    = gradeMatch[1];
  const classNum = classNumMatch[1];
  const marker   = grade + '-' + classNum; // e.g. "3-1"

  // 학년별 컬럼 (1-indexed)
  // J=10, K=11, L=12 / N=14, O=15, P=16 / R=18, S=19, T=20 / V=22, W=23, X=24
  const gradeColMap = {
    '3': {numCol:10, countCol:11, nameCol:12},
    '4': {numCol:14, countCol:15, nameCol:16},
    '5': {numCol:18, countCol:19, nameCol:20},
    '6': {numCol:22, countCol:23, nameCol:24}
  };
  const cols = gradeColMap[grade];
  if (!cols) return;

  // 번호열 전체 읽기
  const numValues   = sheet.getRange(2, cols.numCol,   lastRow-1, 1).getValues();
  const nameValues  = sheet.getRange(2, cols.nameCol,  lastRow-1, 1).getValues();
  const countValues = sheet.getRange(2, cols.countCol, lastRow-1, 1).getValues();

  // marker("3-1") 위치 찾아서 해당 반 row 범위 결정
  let startIdx = -1;
  let endIdx   = numValues.length - 1;

  for (let i = 0; i < numValues.length; i++) {
    const val = String(numValues[i][0]).trim();
    if (val === marker) {
      startIdx = i + 1; // 마커 다음 행부터 학생
    } else if (startIdx > 0 && i >= startIdx) {
      // 숫자가 아닌 값이 나오면 다음 반 시작 → 종료
      if (val !== '' && isNaN(Number(val))) {
        endIdx = i - 1;
        break;
      }
    }
  }
  if (startIdx === -1) return;

  // 모둠장 이름 매칭 후 횟수 +1
  leaders.forEach(leaderName => {
    for (let i = startIdx; i <= endIdx; i++) {
      if (String(nameValues[i][0]).trim() === String(leaderName).trim()) {
        countValues[i][0] = (Number(countValues[i][0]) || 0) + 1;
      }
    }
  });

  sheet.getRange(2, cols.countCol, lastRow-1, 1).setValues(countValues);
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
