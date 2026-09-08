/**
 * ==============================================================================
 * 💑 부부 공유 캘린더 - Google Apps Script 연동 코드 (v2.1 완벽 동기화)
 * ==============================================================================
 * 
 * [1분 설정 방법]
 * 1. 구글 스프레드시트(Google Sheets)를 새로 하나 만듭니다. (제목: 부부 캘린더)
 * 2. 상단 메뉴 [확장 프로그램] -> [Apps Script] 클릭
 * 3. 기존 코드를 모두 지우고 아래 코드를 그대로 붙여넣은 후 저장(💾)
 * 4. 우측 상단 파란색 [배포] -> [새 배포] 클릭
 *    - 유형: [웹 앱] (톱니바퀴 아이콘)
 *    - 설명: 부부 캘린더 실시간 연동
 *    - 다음 사용자로 실행: [나 (내 계정)]
 *    - 액세스 권한이 있는 사용자: [모든 사용자 (Anyone)]  <-- ⭐ 필수!
 * 5. [배포] 클릭 후 발급된 "웹 앱 URL"을 복사하여 부부 캘린더 앱 [⚙️ 설정]에 등록!
 * ==============================================================================
 */

const SCHEDULE_SHEET_NAME = 'Schedules';
const CONFIG_SHEET_NAME = 'Config';

// 일정 시트 초기화 및 가져오기
function getOrCreateScheduleSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SCHEDULE_SHEET_NAME);
    sheet.appendRow(['id', 'title', 'date', 'allDay', 'startTime', 'endTime', 'assignee', 'category', 'completed', 'memo', 'updatedAt']);
    sheet.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#f3e8ff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// 설정/D-Day 시트 초기화 및 가져오기
function getOrCreateConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG_SHEET_NAME);
    sheet.appendRow(['key', 'value', 'updatedAt']);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#fce7f3');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// GET 요청 처리 (일정 및 D-Day 조회)
function doGet(e) {
  try {
    if (e.parameter && e.parameter.data) {
      return handleDataSync(e.parameter.data);
    }

    const schSheet = getOrCreateScheduleSheet();
    const data = schSheet.getDataRange().getValues();
    const schedules = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0]) {
        let dateVal = row[2];
        if (dateVal instanceof Date) {
          dateVal = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        }
        let startTimeVal = row[4];
        if (startTimeVal instanceof Date) {
          startTimeVal = Utilities.formatDate(startTimeVal, Session.getScriptTimeZone(), 'HH:mm');
        }
        let endTimeVal = row[5];
        if (endTimeVal instanceof Date) {
          endTimeVal = Utilities.formatDate(endTimeVal, Session.getScriptTimeZone(), 'HH:mm');
        }

        schedules.push({
          id: String(row[0]),
          title: String(row[1] || ''),
          date: String(dateVal || ''),
          allDay: row[3] === true || row[3] === 'TRUE',
          startTime: String(startTimeVal || ''),
          endTime: String(endTimeVal || ''),
          assignee: String(row[6] || 'couple'),
          category: String(row[7] || '일정'),
          completed: row[8] === true || row[8] === 'TRUE',
          memo: String(row[9] || '')
        });
      }
    }

    // D-Day 설정값 읽기
    const cfgSheet = getOrCreateConfigSheet();
    const cfgData = cfgSheet.getDataRange().getValues();
    let dday = { title: '결혼기념일', date: '2024-05-18' };
    for (let i = 1; i < cfgData.length; i++) {
      if (cfgData[i][0] === 'dday' && cfgData[i][1]) {
        try {
          dday = JSON.parse(cfgData[i][1]);
        } catch(err) {}
      }
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      schedules: schedules,
      dday: dday
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// POST 요청 처리
function doPost(e) {
  try {
    let rawContent = '';
    if (e.postData && e.postData.contents) {
      rawContent = e.postData.contents;
    } else if (e.parameter && e.parameter.data) {
      rawContent = e.parameter.data;
    }

    return handleDataSync(rawContent);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// 시트에 데이터 저장 공통 함수
function handleDataSync(rawJson) {
  const schSheet = getOrCreateScheduleSheet();
  const parsed = JSON.parse(rawJson);
  const schedules = parsed.schedules || [];
  const dday = parsed.dday;
  const updatedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  // 1. 기존 일정 지우고 새로 쓰기
  const lastRow = schSheet.getLastRow();
  if (lastRow > 1) {
    schSheet.deleteRows(2, lastRow - 1);
  }

  if (schedules.length > 0) {
    const rows = schedules.map(s => [
      s.id,
      s.title,
      s.date,
      s.allDay ? 'TRUE' : 'FALSE',
      s.startTime || '',
      s.endTime || '',
      s.assignee || 'couple',
      s.category || '일정',
      s.completed ? 'TRUE' : 'FALSE',
      s.memo || '',
      updatedAt
    ]);
    schSheet.getRange(2, 1, rows.length, 11).setValues(rows);
  }

  // 2. D-Day 설정 저장
  if (dday) {
    const cfgSheet = getOrCreateConfigSheet();
    const cfgLastRow = cfgSheet.getLastRow();
    if (cfgLastRow > 1) {
      cfgSheet.deleteRows(2, cfgLastRow - 1);
    }
    cfgSheet.appendRow(['dday', JSON.stringify(dday), updatedAt]);
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    message: 'Saved successfully',
    count: schedules.length
  })).setMimeType(ContentService.MimeType.JSON);
}
