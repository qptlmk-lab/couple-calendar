/**
 * ==============================================================================
 * 💑 부부 공유 캘린더 - Google Apps Script 연동 코드 (v2 완벽 안정화)
 * ==============================================================================
 * 
 * [1분 설정 방법]
 * 1. 구글 스프레드시트(Google Sheets)를 새로 하나 만듭니다. (제목: 부부 캘린더)
 * 2. 상단 메뉴 [확장 프로그램] -> [Apps Script] 클릭
 * 3. 기존 코드를 모두 지우고 아래 코드를 그대로 붙여넣은 후 저장(💾)
 * 4. 우측 상단 파란색 [배포] -> [새 배포] 클릭
 *    - 유형: [웹 앱] (톱니바퀴 아이콘)
 *    - 설명: 부부 캘린더 연동
 *    - 다음 사용자로 실행: [나 (내 계정)]
 *    - 액세스 권한이 있는 사용자: [모든 사용자 (Anyone)]  <-- ⭐ 필수!
 * 5. [배포] 클릭 후 발급된 "웹 앱 URL"을 복사하여 부부 캘린더 앱에 등록!
 * ==============================================================================
 */

const SHEET_NAME = 'Schedules';

function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['id', 'title', 'date', 'allDay', 'startTime', 'endTime', 'assignee', 'category', 'completed', 'memo', 'updatedAt']);
    sheet.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#f3e8ff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// GET 요청 처리 (일정 조회 및 쿼리 파라미터 저장 지원)
function doGet(e) {
  try {
    // 만약 GET으로 데이터를 넘겼을 경우 처리
    if (e.parameter && e.parameter.data) {
      return handleDataSync(e.parameter.data);
    }

    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
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

    return ContentService.createTextOutput(JSON.stringify({
      status: 'success',
      schedules: schedules
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
  const sheet = getOrCreateSheet();
  const parsed = JSON.parse(rawJson);
  const schedules = parsed.schedules || [];
  const updatedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  // 기존 데이터 지우기
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.deleteRows(2, lastRow - 1);
  }

  // 새 데이터 쓰기
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
    sheet.getRange(2, 1, rows.length, 11).setValues(rows);
  }

  return ContentService.createTextOutput(JSON.stringify({
    status: 'success',
    message: 'Saved successfully',
    count: schedules.length
  })).setMimeType(ContentService.MimeType.JSON);
}
