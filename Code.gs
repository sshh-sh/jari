const SHEET_NAME   = '모둠뽑기';
const SHEET_DETAIL = '모둠뽑기_보기';

// 학년별 명렬표 칸 위치: 반(마커) / 번호 / 이름 / 모둠장 횟수
const GRADE_COLS = {
  '3': {classCol:10, numCol:11, nameCol:12, countCol:13}, // J~M
  '4': {classCol:14, numCol:15, nameCol:16, countCol:17}, // N~Q
  '5': {classCol:18, numCol:19, nameCol:20, countCol:21}, // R~U
  '6': {classCol:22, numCol:23, nameCol:24, countCol:25}  // V~Y
};

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

// ===== 통합 시트 저장 (A~I 결과표. 최신 기록이 맨 위 2행에 옴) =====
function saveDetailSheet(classId, record) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_DETAIL);

  const maxGroups = 6;
  const resultHeaders = ['반', '년도', '월', '1모둠', '2모둠', '3모둠', '4모둠', '5모둠', '6모둠'];

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_DETAIL);
    sheet.getRange(1,1,1,9).setValues([resultHeaders]);
    sheet.getRange(1,1,1,9).setFontWeight('bold').setBackground('#534AB7').setFontColor('white');
    sheet.setFrozenRows(1);
    for (let i = 4; i <= 9; i++) sheet.setColumnWidth(i, 180);
  }

  const className = record.className || classId;

  // 같은 반+월 기존 행 삭제 (A~I 열만, 명렬표 칸은 건드리지 않음)
  const resultLastRow = getResultLastRow(sheet);
  if (resultLastRow > 1) {
    const existing = sheet.getRange(2, 1, resultLastRow-1, 9).getValues();
    for (let i = existing.length-1; i >= 0; i--) {
      if (existing[i][0] === className && existing[i][1] === record.year && existing[i][2] === record.month) {
        deleteResultRow(sheet, i+2);
      }
    }
  }

  const row = [className, record.year, record.month];
  for (let i = 0; i < maxGroups; i++) {
    const group = record.groups[i];
    row.push(group ? group.students.join(', ') : '');
  }
  insertResultRowAtTop(sheet, row);

  if (record.leaders && record.leaders.length > 0) {
    updateLeaderCount(sheet, record.leaders, className);
  }
}

// 결과표(A~I)에서 실제로 데이터가 채워진 마지막 행 (명렬표가 더 아래까지 있어도 무시)
function getResultLastRow(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 1;
  const colA = sheet.getRange(2, 1, lastRow-1, 1).getValues();
  let last = 1;
  for (let i = 0; i < colA.length; i++) {
    if (colA[i][0] !== '') last = i+2;
  }
  return last;
}

// 결과표(A~I)의 2행(헤더 바로 아래)에 새 행 삽입 — 최신순 정렬. J열 이후(명렬표)는 그대로 둠
function insertResultRowAtTop(sheet, row) {
  const resultLastRow = getResultLastRow(sheet);
  if (resultLastRow > 1) {
    const oldValues = sheet.getRange(2, 1, resultLastRow-1, 9).getValues();
    sheet.getRange(3, 1, resultLastRow-1, 9).setValues(oldValues);
  }
  sheet.getRange(2, 1, 1, 9).setValues([row]);
  restripeResultRows(sheet);
}

// 결과표(A~I)에서 특정 행 삭제(아래 행들을 위로 당김). J열 이후(명렬표)는 그대로 둠
function deleteResultRow(sheet, rowIndex) {
  const resultLastRow = getResultLastRow(sheet);
  if (rowIndex < resultLastRow) {
    const below = sheet.getRange(rowIndex+1, 1, resultLastRow-rowIndex, 9).getValues();
    sheet.getRange(rowIndex, 1, resultLastRow-rowIndex, 9).setValues(below);
  }
  sheet.getRange(resultLastRow, 1, 1, 9).clearContent();
}

function restripeResultRows(sheet) {
  const resultLastRow = getResultLastRow(sheet);
  for (let r = 2; r <= resultLastRow; r++) {
    sheet.getRange(r,1,1,9).setBackground((r%2===0) ? '#f0effe' : '#ffffff');
  }
}

// ===== 모둠장 횟수 업데이트 =====
function updateLeaderCount(sheet, leaders, className) {
  const gradeMatch    = className.match(/(\d+)학년/);
  const classNumMatch = className.match(/(\d+)반/);
  if (!gradeMatch || !classNumMatch) return;
  const grade  = gradeMatch[1];
  const marker = grade + '-' + classNumMatch[1];

  const cols = GRADE_COLS[grade];
  if (!cols) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const classValues = sheet.getRange(2, cols.classCol, lastRow-1, 1).getValues();
  const nameValues  = sheet.getRange(2, cols.nameCol,  lastRow-1, 1).getValues();
  const countValues = sheet.getRange(2, cols.countCol, lastRow-1, 1).getValues();

  let startIdx = -1, endIdx = classValues.length - 1;
  for (let i = 0; i < classValues.length; i++) {
    const val = String(classValues[i][0]).trim();
    if (val === marker) { startIdx = i; }
    else if (startIdx >= 0 && i > startIdx && val !== '') { endIdx = i-1; break; }
  }
  if (startIdx === -1) return;

  leaders.forEach(leaderName => {
    for (let i = startIdx; i <= endIdx; i++) {
      if (String(nameValues[i][0]).trim() === String(leaderName).trim()) {
        countValues[i][0] = (Number(countValues[i][0]) || 0) + 1;
      }
    }
  });

  sheet.getRange(2, cols.countCol, lastRow-1, 1).setValues(countValues);
}

// ===================================================================
// ★ 1회용 설정 함수 - Apps Script 편집기에서 한 번 실행하면 K열부터 학년별
// 명렬표 칸(반/번호/이름/모둠장 횟수)이 자동으로 생성됩니다.
// 지금까지 저장된 모둠 결과에 등장한 학생들을 모아 만들며, 이미 채워진
// 모둠장 횟수는 그대로 보존됩니다. 이후 새 학생이 추가되면 다시 실행해도 됨.
// ===================================================================
function setupLeaderRoster() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DETAIL);
  if (!sheet) { Logger.log('모둠뽑기_보기 시트가 없습니다. 먼저 결과를 한 번 저장해 주세요.'); return; }

  const resultLastRow = getResultLastRow(sheet);
  if (resultLastRow < 2) { Logger.log('저장된 결과가 없습니다.'); return; }

  const data = sheet.getRange(2, 1, resultLastRow-1, 9).getValues();
  const rosterByClass = {}; // { "3-1": Set(이름) }

  data.forEach(row => {
    const className = String(row[0]);
    const gradeMatch = className.match(/(\d+)학년/);
    const classNumMatch = className.match(/(\d+)반/);
    if (!gradeMatch || !classNumMatch) return;
    const marker = gradeMatch[1] + '-' + classNumMatch[1];
    if (!rosterByClass[marker]) rosterByClass[marker] = new Set();
    for (let c = 3; c <= 8; c++) {
      const cell = row[c];
      if (!cell) continue;
      String(cell).split(',').map(s => s.trim()).filter(Boolean).forEach(name => rosterByClass[marker].add(name));
    }
  });

  const markersByGrade = {};
  Object.keys(rosterByClass).forEach(marker => {
    const grade = marker.split('-')[0];
    if (!markersByGrade[grade]) markersByGrade[grade] = [];
    markersByGrade[grade].push(marker);
  });

  Object.keys(markersByGrade).forEach(grade => {
    const cols = GRADE_COLS[grade];
    if (!cols) return;
    markersByGrade[grade].sort();
    writeGradeRoster(sheet, cols, markersByGrade[grade], rosterByClass);
  });

  Logger.log('명렬표 생성 완료!');
}

function writeGradeRoster(sheet, cols, markers, rosterByClass) {
  const existingCount = readExistingLeaderCounts(sheet, cols);

  sheet.getRange(1, cols.classCol, 1, 4).setValues([['반','번호','이름','모둠장 횟수']]);
  sheet.getRange(1, cols.classCol, 1, 4).setFontWeight('bold').setBackground('#534AB7').setFontColor('white');

  let row = 2;
  markers.forEach(marker => {
    const names = [...rosterByClass[marker]].sort();
    names.forEach((name, idx) => {
      sheet.getRange(row, cols.classCol).setValue(idx === 0 ? marker : '');
      sheet.getRange(row, cols.numCol).setValue(idx + 1);
      sheet.getRange(row, cols.nameCol).setValue(name);
      sheet.getRange(row, cols.countCol).setValue(existingCount[marker + '|' + name] || 0);
      row++;
    });
  });
}

// 이미 명렬표에 있던 학생의 모둠장 횟수를 보존하기 위해 미리 읽어둠
function readExistingLeaderCounts(sheet, cols) {
  const lastRow = sheet.getLastRow();
  const result = {};
  if (lastRow < 2) return result;

  const classValues = sheet.getRange(2, cols.classCol, lastRow-1, 1).getValues();
  const nameValues  = sheet.getRange(2, cols.nameCol,  lastRow-1, 1).getValues();
  const countValues = sheet.getRange(2, cols.countCol, lastRow-1, 1).getValues();

  let curMarker = '';
  for (let i = 0; i < classValues.length; i++) {
    const marker = String(classValues[i][0]).trim();
    if (marker) curMarker = marker;
    const name = String(nameValues[i][0]).trim();
    if (curMarker && name) result[curMarker + '|' + name] = Number(countValues[i][0]) || 0;
  }
  return result;
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
