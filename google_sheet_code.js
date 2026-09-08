/**
 * ==============================================================================
 * 💑 부부 공유 캘린더 - Google Apps Script 완벽 연동 코드 (v3.0)
 * ==============================================================================
 * 
 * [1분 설정 핵심 체크리스트]
 * 1. 구글 스프레드시트 -> [확장 프로그램] -> [Apps Script]
 * 2. 기존 코드를 모두 지우고 이 코드 전체를 붙여넣은 후 저장(💾)
 * 3. 오른쪽 위 [배포] -> [새 배포] 클릭
 *    - 유형: [웹 앱] 선택
 *    - 설명: 부부 캘린더
 *    - 다음 사용자로 실행: [나]
 *    - 액세스 권한: [모든 사용자 (Anyone)] ⭐ 필수! (로그인 없이 동기화)
 * 4. [배포] 누르고 발급된 웹 앱 URL(/exec로 끝나는 주소) 복사 후 앱에 등록!
 * ==============================================================================
 */

const SCHEDULE_SHEET_NAME = 'Schedules';
const CONFIG_SHEET_NAME = 'Config';

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

// GET 요청 처리 (데이터 조회 및 CORS 없는 저장 지원)
function doGet(e) {
  try {
    // 1. URL 파라미터로 저장 요청이 들어온 경우 (CORS 제약 완전 해결)
    if (e.parameter && e.parameter.data) {
      return handleDataSync(e.parameter.data, e.parameter.callback);
    }

    // 2. 헬스체크/연동 테스트
    if (e.parameter && e.parameter.action === 'test') {
      return respondJson({ status: 'success', message: '연동 성공! 구글 시트와 정상 연결되었습니다.' }, e.parameter.callback);
    }

    // 3. 일반 일정 조회
    const schSheet = getOrCreateScheduleSheet();
    const data = schSheet.getDataRange().getValues();
    const schedules = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0]) {
        let dateVal = row[2] instanceof Date ? Utilities.formatDate(row[2], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(row[2] || '');
        let startTimeVal = row[4] instanceof Date ? Utilities.formatDate(row[4], Session.getScriptTimeZone(), 'HH:mm') : String(row[4] || '');
        let endTimeVal = row[5] instanceof Date ? Utilities.formatDate(row[5], Session.getScriptTimeZone(), 'HH:mm') : String(row[5] || '');

        schedules.push({
          id: String(row[0]),
          title: String(row[1] || ''),
          date: dateVal,
          allDay: row[3] === true || String(row[3]).toUpperCase() === 'TRUE',
          startTime: startTimeVal,
          endTime: endTimeVal,
          assignee: String(row[6] || 'couple'),
          category: String(row[7] || '일정'),
          completed: row[8] === true || String(row[8]).toUpperCase() === 'TRUE',
          memo: String(row[9] || '')
        });
      }
    }

    // D-Day 설정 조회
    const cfgSheet = getOrCreateConfigSheet();
    const cfgData = cfgSheet.getDataRange().getValues();
    let dday = { title: '결혼기념일', date: '2024-05-18' };
    for (let i = 1; i < cfgData.length; i++) {
      if (cfgData[i][0] === 'dday' && cfgData[i][1]) {
        try { dday = JSON.parse(cfgData[i][1]); } catch(err) {}
      }
    }

    return respondJson({
      status: 'success',
      schedules: schedules,
      dday: dday
    }, e.parameter ? e.parameter.callback : null);

  } catch (err) {
    return respondJson({
      status: 'error',
      message: err.toString()
    }, e.parameter ? e.parameter.callback : null);
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

    return handleDataSync(rawContent, null);

  } catch (err) {
    return respondJson({
      status: 'error',
      message: err.toString()
    });
  }
}

// 데이터 저장 핸들러
function handleDataSync(rawJson, callback) {
  const schSheet = getOrCreateScheduleSheet();
  const parsed = typeof rawJson === 'string' ? JSON.parse(rawJson) : rawJson;
  const schedules = parsed.schedules || [];
  const dday = parsed.dday;
  const updatedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  // 기존 일정 지우기
  const lastRow = schSheet.getLastRow();
  if (lastRow > 1) {
    schSheet.deleteRows(2, lastRow - 1);
  }

  // 새 일정 쓰기
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

  // D-Day 저장
  if (dday) {
    const cfgSheet = getOrCreateConfigSheet();
    const cfgLastRow = cfgSheet.getLastRow();
    if (cfgLastRow > 1) {
      cfgSheet.deleteRows(2, cfgLastRow - 1);
    }
    cfgSheet.appendRow(['dday', JSON.stringify(dday), updatedAt]);
  }

  return respondJson({
    status: 'success',
    message: 'Saved successfully',
    count: schedules.length
  }, callback);
}

// JSON / JSONP 공통 응답 생성
function respondJson(obj, callback) {
  const jsonStr = JSON.stringify(obj);
  if (callback) {
    return ContentService.createTextOutput(`${callback}(${jsonStr})`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(jsonStr)
    .setMimeType(ContentService.MimeType.JSON);
}
