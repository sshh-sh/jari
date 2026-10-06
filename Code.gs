const SHEET_NAME   = '모둠뽑기';
const SHEET_DETAIL = '모둠뽑기_보기';

// 학년별 명렬표 칸 위치: 반(마커) / 번호 / 이름 / 모둠장 기록(월) / 매니저
const GRADE_COLS = {
  '3': {classCol:10, numCol:11, nameCol:12, countCol:13, mgrCol:14}, // J~N
  '4': {classCol:15, numCol:16, nameCol:17, countCol:18, mgrCol:19}, // O~S
  '5': {classCol:20, numCol:21, nameCol:22, countCol:23, mgrCol:24}, // T~X
  '6': {classCol:25, numCol:26, nameCol:27, countCol:28, mgrCol:29}  // Y~AC
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
    else if (action === 'bulkSetRoster') result = bulkSetRoster(payload.marker, payload.records);
    else if (action === 'createClassRoster') result = createClassRoster(payload.marker, payload.records);
    else if (action === 'backup') result = saveBackup(payload);
    else result = {success: false, message: '알 수 없는 액션'};

    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({success: false, message: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  const action = e.parameter && e.parameter.action;
  const callback = e.parameter && e.parameter.callback;

  function respond(obj) {
    const json = JSON.stringify(obj);
    if (callback) {
      return ContentService.createTextOutput(callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'setupRoster') {
    setupLeaderRoster();
    return respond({status:'ok', message:'명렬표 생성 완료'});
  }
  if (action === 'addManagerColumns') {
    const result = addManagerColumns();
    return respond({status:'ok', message: result});
  }
  if (action === 'roster') {
    const marker = e.parameter.marker || '';
    return respond({status:'ok', students: getRosterInfo(marker)});
  }
  if (action === 'history') {
    return respond({status:'ok', records: getAllHistory()});
  }
  if (action === 'backup') {
    return respond({status:'ok', backup: getLatestBackup()});
  }
  return respond({status:'ok', message:'모둠뽑기 GAS 작동 중'});
}

// ===== 전체 기록 내려주기 (새 컴퓨터에서 기록 복원용) =====
// '모둠뽑기' 시트의 모든 기록을 반 이름과 함께 돌려준다.
// 반 이름은 '모둠뽑기_보기' 시트에서 (년도·월·모둠별 학생)이 똑같은 행을 찾아 알아낸다.
function getAllHistory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const main = ss.getSheetByName(SHEET_NAME);
  if (!main) return [];
  const rows = main.getDataRange().getValues();

  // 보기 시트: 서명(년|월|1모둠|2모둠|...) → 반 이름
  const nameBySig = {};
  const detail = ss.getSheetByName(SHEET_DETAIL);
  if (detail) {
    const last = getResultLastRow(detail);
    if (last > 1) {
      detail.getRange(2, 1, last-1, 9).getValues().forEach(function(r) {
        if (r[0] === '') return;
        nameBySig[[r[1], r[2], r[3], r[4], r[5], r[6], r[7], r[8]].join('|')] = String(r[0]);
      });
    }
  }

  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (r[0] === '' || r[4] === '') continue;
    let groups;
    try { groups = JSON.parse(r[4]); } catch (err) { continue; }
    if (!Array.isArray(groups)) continue;
    const parts = [r[1], r[2]];
    for (let g = 0; g < 6; g++) parts.push(groups[g] && groups[g].students ? groups[g].students.join(', ') : '');
    const date = (r[3] instanceof Date)
      ? Utilities.formatDate(r[3], 'Asia/Seoul', 'yyyy. M. d.')
      : String(r[3]);
    out.push({
      classId: String(r[0]),
      className: nameBySig[parts.join('|')] || '',
      year: Number(r[1]),
      month: Number(r[2]),
      date: date,
      groups: groups
    });
  }
  return out;
}

// ===== 앱 전체 백업 (다른 컴퓨터로 옮기기용) =====
// '모둠뽑기_백업' 시트에 [백업ID, 저장시각, 요약, 조각번호, 조각수, 내용] 행으로 저장한다.
// 셀 하나에 5만 자까지만 들어가서 내용을 조각내 여러 행에 나눠 담고, 최근 백업 5개만 남긴다.
const SHEET_BACKUP = '모둠뽑기_백업';
const BACKUP_CHUNK = 40000;
const BACKUP_KEEP  = 5;

function saveBackup(payload) {
  if (!payload || !payload.id || !payload.state) return {success: false, message: '백업 내용 없음'};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_BACKUP);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_BACKUP);
    sheet.getRange(1,1,1,6).setValues([['백업ID','저장시각','요약','조각번호','조각수','내용']]);
    sheet.getRange(1,1,1,6).setFontWeight('bold').setBackground('#534AB7').setFontColor('white');
    sheet.setFrozenRows(1);
  }
  const json = JSON.stringify({id: payload.id, savedAt: payload.savedAt, summary: payload.summary, state: payload.state});
  const total = Math.ceil(json.length / BACKUP_CHUNK);
  const rows = [];
  for (let i = 0; i < total; i++) {
    // 조각이 '='·숫자 등으로 시작하면 시트가 수식/숫자로 바꿔 버리므로 앞에 '§'를 붙여 항상 글자로 저장
    rows.push([String(payload.id), String(payload.savedAt || ''), String(payload.summary || ''), i + 1, total,
               '§' + json.slice(i * BACKUP_CHUNK, (i + 1) * BACKUP_CHUNK)]);
  }
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);

  // 오래된 백업 정리 (최근 BACKUP_KEEP개만)
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().map(r => String(r[0]));
  const ids = [];
  data.forEach(id => { if (ids.indexOf(id) < 0) ids.push(id); });
  const removeIds = ids.slice(0, Math.max(0, ids.length - BACKUP_KEEP));
  for (let i = data.length - 1; i >= 0; i--) {
    if (removeIds.indexOf(data[i]) >= 0) sheet.deleteRow(i + 2);
  }
  return {success: true, message: '백업 저장 완료'};
}

function getLatestBackup() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_BACKUP);
  if (!sheet || sheet.getLastRow() < 2) return null;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  const lastId = String(rows[rows.length - 1][0]);
  const parts = rows.filter(r => String(r[0]) === lastId).sort((a, b) => a[3] - b[3]);
  if (!parts.length || parts.length !== Number(parts[0][4])) return null;
  try {
    return JSON.parse(parts.map(r => String(r[5]).replace(/^§/, '')).join(''));
  } catch (err) {
    return null;
  }
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
    updateLeaderCount(sheet, record.leaders, className, record.month);
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

// ===== 모둠장 기록(월) 업데이트 =====
// 명렬표에 반/학생이 아직 없으면(새 반, 새 학생) 명렬표를 다시 만든 뒤 한 번 더 시도한다.
function updateLeaderCount(sheet, leaders, className, month) {
  const gradeMatch    = className.match(/(\d+)학년/);
  const classNumMatch = className.match(/(\d+)반/);
  if (!gradeMatch || !classNumMatch) return;
  const grade  = gradeMatch[1];
  const marker = grade + '-' + classNumMatch[1];

  const cols = GRADE_COLS[grade];
  if (!cols) return;

  if (!applyLeaderIncrement(sheet, cols, marker, leaders, month)) {
    setupLeaderRoster();
    applyLeaderIncrement(sheet, cols, marker, leaders, month);
  }
}

// marker 반의 leaders 각각의 기록 칸에 "N월"을 추가(같은 달 중복 추가 안 함).
// leaders 전원이 명렬표에서 매칭되면 true, 하나라도 못 찾으면 false.
function applyLeaderIncrement(sheet, cols, marker, leaders, month) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return false;

  const classValues  = sheet.getRange(2, cols.classCol, lastRow-1, 1).getValues();
  const nameValues   = sheet.getRange(2, cols.nameCol,  lastRow-1, 1).getValues();
  const recordValues = sheet.getRange(2, cols.countCol, lastRow-1, 1).getValues();

  let startIdx = -1, endIdx = classValues.length - 1;
  for (let i = 0; i < classValues.length; i++) {
    const val = String(classValues[i][0]).trim();
    if (val === marker) { startIdx = i; }
    else if (startIdx >= 0 && i > startIdx && val !== '') { endIdx = i-1; break; }
  }
  if (startIdx === -1) return false;

  const token = month + '월';
  let allMatched = true;
  leaders.forEach(leaderName => {
    let matched = false;
    for (let i = startIdx; i <= endIdx; i++) {
      if (String(nameValues[i][0]).trim() === String(leaderName).trim()) {
        const months = String(recordValues[i][0] || '').trim().split(/\s+/).filter(Boolean);
        if (!months.includes(token)) months.push(token);
        recordValues[i][0] = months.join(' ');
        matched = true;
      }
    }
    if (!matched) allMatched = false;
  });

  sheet.getRange(2, cols.countCol, lastRow-1, 1).setValues(recordValues);
  return allMatched;
}

// ===================================================================
// ★ 1회용 마이그레이션 - 학년별 명렬표 블록마다 '매니저' 열을 하나씩 삽입한다.
// 실제 시트 열 삽입(insertColumnBefore)을 쓰므로 기존 데이터는 자동으로
// 오른쪽으로 밀려서 보존된다. 이미 삽입되어 있으면(헤더가 '매니저'면) 다시 실행하지 않음.
// ===================================================================
function addManagerColumns() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DETAIL);
  if (!sheet) return '모둠뽑기_보기 시트가 없습니다.';

  const grades = ['3','4','5','6'];
  const maxNeeded = Math.max.apply(null, grades.map(g => GRADE_COLS[g].mgrCol));
  if (sheet.getMaxColumns() < maxNeeded) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), maxNeeded - sheet.getMaxColumns());
  }

  let inserted = 0;
  grades.forEach(grade => {
    const cols = GRADE_COLS[grade];
    const header = String(sheet.getRange(1, cols.mgrCol).getValue()).trim();
    if (header === '매니저') return; // 이미 마이그레이션됨
    sheet.insertColumnBefore(cols.mgrCol);
    sheet.getRange(1, cols.mgrCol).setValue('매니저')
      .setFontWeight('bold').setBackground('#534AB7').setFontColor('white');
    inserted++;
  });
  return inserted > 0 ? (inserted + '개 학년 블록에 매니저 열 삽입 완료') : '이미 모든 학년에 매니저 열이 있습니다.';
}

// marker(예: "3-1") 반의 명렬표를 읽어 이름별 모둠장 횟수/매니저 여부를 반환.
// 모둠장 기록은 쉼표 없이 띄어쓰기로만 구분되어 있음("4월 5월 7월").
function getRosterInfo(marker) {
  const grade = marker.split('-')[0];
  const cols = GRADE_COLS[grade];
  if (!cols) return [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DETAIL);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const classValues  = sheet.getRange(2, cols.classCol, lastRow-1, 1).getValues();
  const nameValues   = sheet.getRange(2, cols.nameCol,  lastRow-1, 1).getValues();
  const recordValues = sheet.getRange(2, cols.countCol, lastRow-1, 1).getValues();
  const mgrValues    = sheet.getRange(2, cols.mgrCol,   lastRow-1, 1).getValues();

  let startIdx = -1, endIdx = classValues.length - 1;
  for (let i = 0; i < classValues.length; i++) {
    const val = String(classValues[i][0]).trim();
    if (val === marker) { startIdx = i; }
    else if (startIdx >= 0 && i > startIdx && val !== '') { endIdx = i-1; break; }
  }
  if (startIdx === -1) return [];

  const result = [];
  for (let i = startIdx; i <= endIdx; i++) {
    const name = String(nameValues[i][0]).trim();
    if (!name) continue;
    const months = String(recordValues[i][0] || '').trim().split(/\s+/).filter(Boolean);
    const isManager = String(mgrValues[i][0] || '').trim() !== '';
    result.push({name, count: months.length, isManager});
  }
  return result;
}

// ===================================================================
// ★ 1회용 설정 함수 - Apps Script 편집기에서 한 번 실행하면 K열부터 학년별
// 명렬표 칸(반/번호/이름/모둠장 기록)이 자동으로 생성됩니다.
// 지금까지 저장된 모둠 결과에 등장한 학생들을 모아 만들며, 이미 채워진
// 모둠장 기록은 그대로 보존됩니다. 이후 새 학생이 추가되면 다시 실행해도 됨.
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
  const existingRecord = readExistingLeaderRecords(sheet, cols);

  sheet.getRange(1, cols.classCol, 1, 5).setValues([['반','번호','이름','모둠장 기록(월)','매니저']]);
  sheet.getRange(1, cols.classCol, 1, 5).setFontWeight('bold').setBackground('#534AB7').setFontColor('white');

  let row = 2;
  markers.forEach(marker => {
    const names = [...rosterByClass[marker]].sort();
    names.forEach((name, idx) => {
      const existing = existingRecord[marker + '|' + name] || {};
      // 반 마커("3-1" 등)를 그대로 두면 구글시트가 날짜로 자동 변환해버리므로 텍스트 서식 강제
      sheet.getRange(row, cols.classCol).setNumberFormat('@').setValue(idx === 0 ? marker : '');
      sheet.getRange(row, cols.numCol).setValue(idx + 1);
      sheet.getRange(row, cols.nameCol).setNumberFormat('@').setValue(name);
      sheet.getRange(row, cols.countCol).setNumberFormat('@').setValue(existing.count || '');
      sheet.getRange(row, cols.mgrCol).setNumberFormat('@').setValue(existing.mgr || '');
      row++;
    });
  });
}

// 이미 명렬표에 있던 학생의 모둠장 기록(월)/매니저 표시를 보존하기 위해 미리 읽어둠
function readExistingLeaderRecords(sheet, cols) {
  const lastRow = sheet.getLastRow();
  const result = {};
  if (lastRow < 2) return result;

  const classValues  = sheet.getRange(2, cols.classCol, lastRow-1, 1).getValues();
  const nameValues   = sheet.getRange(2, cols.nameCol,  lastRow-1, 1).getValues();
  const recordValues = sheet.getRange(2, cols.countCol, lastRow-1, 1).getValues();
  const mgrValues    = sheet.getRange(2, cols.mgrCol,   lastRow-1, 1).getValues();

  let curMarker = '';
  for (let i = 0; i < classValues.length; i++) {
    const marker = String(classValues[i][0]).trim();
    if (marker) curMarker = marker;
    const name = String(nameValues[i][0]).trim();
    if (curMarker && name) {
      result[curMarker + '|' + name] = {
        count: String(recordValues[i][0] || '').trim(),
        mgr: String(mgrValues[i][0] || '').trim()
      };
    }
  }
  return result;
}

// 캡처 복구용: marker 반의 학생별 모둠장 기록(월 배열)/매니저 여부를 한 번에 직접 설정(덮어쓰기).
// records: [{name, months:[3,5,6], isManager:true}]
function bulkSetRoster(marker, records) {
  const grade = marker.split('-')[0];
  const cols = GRADE_COLS[grade];
  if (!cols) return {success:false, message:'학년 매핑을 찾을 수 없습니다: ' + marker};

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DETAIL);
  if (!sheet) return {success:false, message:SHEET_DETAIL + ' 시트가 없습니다.'};

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return {success:false, message:'명렬표가 비어 있습니다.'};

  const classValues = sheet.getRange(2, cols.classCol, lastRow-1, 1).getValues();
  const nameValues  = sheet.getRange(2, cols.nameCol,  lastRow-1, 1).getValues();

  let startIdx = -1, endIdx = classValues.length - 1;
  for (let i = 0; i < classValues.length; i++) {
    const val = String(classValues[i][0]).trim();
    if (val === marker) { startIdx = i; }
    else if (startIdx >= 0 && i > startIdx && val !== '') { endIdx = i-1; break; }
  }
  if (startIdx === -1) return {success:false, message:marker + ' 반을 명렬표에서 찾을 수 없습니다.'};

  const countRange = sheet.getRange(2, cols.countCol, lastRow-1, 1);
  const mgrRange   = sheet.getRange(2, cols.mgrCol,   lastRow-1, 1);
  const countVals  = countRange.getValues();
  const mgrVals    = mgrRange.getValues();

  const unmatched = [];
  records.forEach(r => {
    let found = false;
    for (let i = startIdx; i <= endIdx; i++) {
      if (String(nameValues[i][0]).trim() === String(r.name).trim()) {
        countVals[i][0] = (r.months && r.months.length) ? r.months.map(m => m + '월').join(' ') : '';
        mgrVals[i][0] = r.isManager ? '매니저' : '';
        found = true;
      }
    }
    if (!found) unmatched.push(r.name);
  });

  countRange.setNumberFormat('@').setValues(countVals);
  mgrRange.setNumberFormat('@').setValues(mgrVals);
  return {success:true, unmatched};
}

// 캡처 복구용: 해당 학년 블록에 marker 반이 명렬표에 아직 없을 때, 새 반을 통째로 추가.
// records: [{name, months:[3,5,6], isManager:true}] (번호는 배열 순서대로 1부터 자동 부여)
function createClassRoster(marker, records) {
  const grade = marker.split('-')[0];
  const cols = GRADE_COLS[grade];
  if (!cols) return {success:false, message:'학년 매핑을 찾을 수 없습니다: ' + marker};

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_DETAIL);
  if (!sheet) return {success:false, message:SHEET_DETAIL + ' 시트가 없습니다.'};

  const lastRow = sheet.getLastRow();
  let insertAt = 2;
  if (lastRow >= 2) {
    const classValues = sheet.getRange(2, cols.classCol, lastRow-1, 1).getValues();
    const nameValues  = sheet.getRange(2, cols.nameCol,  lastRow-1, 1).getValues();
    for (let i = 0; i < classValues.length; i++) {
      const marker2 = String(classValues[i][0]).trim();
      if (marker2 === marker) return {success:false, message:marker + ' 반이 이미 명렬표에 있습니다. bulkSetRoster를 사용하세요.'};
      if (marker2 || String(nameValues[i][0]).trim()) insertAt = 2 + i + 1;
    }
  }

  records.forEach((r, idx) => {
    const row = insertAt + idx;
    sheet.getRange(row, cols.classCol).setNumberFormat('@').setValue(idx === 0 ? marker : '');
    sheet.getRange(row, cols.numCol).setValue(idx + 1);
    sheet.getRange(row, cols.nameCol).setNumberFormat('@').setValue(r.name);
    sheet.getRange(row, cols.countCol).setNumberFormat('@').setValue((r.months && r.months.length) ? r.months.map(m => m + '월').join(' ') : '');
    sheet.getRange(row, cols.mgrCol).setNumberFormat('@').setValue(r.isManager ? '매니저' : '');
  });

  return {success:true, added: records.length};
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
