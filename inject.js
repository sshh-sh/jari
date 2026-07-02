monaco.editor.getModels()[0].setValue(`/**
 * 관피타 출첵 - Google Apps Script 백엔드
 * - doGet  ?action=get&key=...  /  ?action=set&key=...&value=...
 * - doPost { app:'kao', action:'set', key:'kao-state', value }
 */

const KAO_SHEET = '오케';

// ===================================================================
// 진입점 - doPost
// ===================================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    let result = {ok: false};

    if (data.app === 'kao' && data.action === 'set' && data.key === 'kao-state') {
      try { writeKaoSheet(JSON.parse(data.value)); } catch(e) {}
      result = {ok: true};
    }

    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, message: err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

// ===================================================================
// 진입점 - doGet (관피타 key-value 저장소)
// ===================================================================
function doGet(e) {
  const action = e.parameter && e.parameter.action;
  const key    = e.parameter && e.parameter.key;
  const value  = e.parameter && e.parameter.value;

  if (action === 'get') {
    const val = kaoGet(key);
    return ContentService.createTextOutput(JSON.stringify({ok: val !== null, value: val})).setMimeType(ContentService.MimeType.JSON);
  }

  if (action === 'set') {
    if (key === 'kao-state') {
      try { writeKaoSheet(JSON.parse(value)); } catch(e) {}
    }
    return ContentService.createTextOutput(JSON.stringify({ok: true})).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput(JSON.stringify({status:'ok', message:'관피타 출첵 GAS 작동 중'})).setMimeType(ContentService.MimeType.JSON);
}

// ===================================================================
// 오케 탭에서 kao-state JSON 읽어오기
// ===================================================================
function kaoGet(key) {
  const sheet = getOrCreateKaoSheet();
  // A1 셀에 key, B1 셀에 JSON 저장
  const rows = sheet.getDataRange().getValues();
  for (let i = 0; i < rows.length; i++) {
    if (rows[i][0] === key) return rows[i][1];
  }
  return null;
}

// ===================================================================
// 오케 탭에 날짜×멤버 출석표 + JSON 백업 저장
// ===================================================================
function writeKaoSheet(state) {
  const sheet = getOrCreateKaoSheet();
  sheet.clearContents();
  sheet.clearFormats();

  const members = state.members || [];
  const dates   = (state.dates  || []).sort();
  const att     = state.attendance || {};
  const labels  = state.labels || {};

  // 1행: JSON 백업 (get용) — 숨겨두기
  sheet.getRange(1, 1, 1, 2).setValues([['kao-state', JSON.stringify(state)]]);
  sheet.setRowHeight(1, 3);
  sheet.getRange(1, 1, 1, 2).setFontColor('#ffffff');

  // 2행: 헤더 — A:날짜, B:행사, C~:멤버
  const header = ['날짜', '행사', ...members];
  sheet.getRange(2, 1, 1, header.length).setValues([header]);
  sheet.getRange(2, 1, 1, header.length).setFontWeight('bold').setBackground('#1D9E75').setFontColor('white');
  sheet.setFrozenRows(2);

  if (!dates.length) return;

  // 날짜별 행: A=날짜, B=행사명(없으면 빈칸), C~=출석
  const rows = dates.map(d => {
    const dayAtt = att[d] || {};
    return [d, labels[d] || '', ...members.map(m => {
      const v = dayAtt[m];
      return v === 'o' ? '참여' : v === 'x' ? '불참' : '-';
    })];
  });

  sheet.getRange(3, 1, rows.length, header.length).setValues(rows);

  // A열 날짜 왼쪽 정렬
  sheet.getRange(2, 1, rows.length + 1, 1).setHorizontalAlignment('left');

  // 참여/불참 색상 (C열부터)
  for (let r = 0; r < rows.length; r++) {
    for (let c = 2; c < header.length; c++) {
      const val  = rows[r][c];
      const cell = sheet.getRange(r + 3, c + 1);
      if (val === '참여')      cell.setBackground('#E1F5EE').setFontColor('#0F6E56');
      else if (val === '불참') cell.setBackground('#FCEBEB').setFontColor('#A32D2D');
      else                     cell.setBackground('#ffffff').setFontColor('#aaaaaa');
    }
  }

  sheet.autoResizeColumns(1, header.length);
  sheet.setColumnWidth(2, 80);
  for (let c = 3; c <= header.length; c++) sheet.setColumnWidth(c, 70);
}

function getOrCreateKaoSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(KAO_SHEET);
  if (!sheet) sheet = ss.insertSheet(KAO_SHEET);
  return sheet;
}
`)
