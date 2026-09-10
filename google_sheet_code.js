/**
 * ==============================================================================
 * 💑 부부 공유 캘린더 - Google Apps Script 실시간 동기화 코드
 * ==============================================================================
 * 
 * [1분 설정 방법]
 * 1. 구글 스프레드시트 -> 상단 메뉴 [확장 프로그램] -> [Apps Script] 클릭
 * 2. 기존 코드를 모두 지우고 이 파일 전체 내용을 붙여넣은 후 저장(💾)
 * 3. 오른쪽 상단 파란색 [배포] -> [새 배포] 클릭
 *    - 유형 선택: [웹 앱] (톱니바퀴 아이콘 클릭)
 *    - 설명: 부부 캘린더 연동
 *    - 다음 사용자로 실행: [나]
 *    - 액세스 권한: [모든 사용자 (Anyone)] ⭐ 필수!
 * 4. [배포] 버튼을 누르고 발급된 웹 앱 URL(/exec 로 끝나는 주소) 복사
 * 5. 달력 화면의 [⚙️] 설정 버튼을 눌러 URL을 입력하고 저장하면 끝!
 * ==============================================================================
 */

const SCHEDULE_SHEET_NAME = 'Schedules';

function getOrCreateScheduleSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SCHEDULE_SHEET_NAME);
    sheet.appendRow(['id', 'assignee', 'title', 'date', 'allDay', 'startTime', 'endTime', 'memo', 'updatedAt']);
    sheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#eff6ff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// GET 요청 처리 (일정 조회 및 쿼리 파라미터 백업 저장 지원)
function doGet(e) {
  try {
    // 1. URL 파라미터로 저장 데이터가 전달된 경우
    if (e.parameter && e.parameter.data) {
      return handleDataSync(e.parameter.data, e.parameter.callback);
    }

    // 2. 연동 테스트 (헬스체크)
    if (e.parameter && e.parameter.action === 'test') {
      return respondJson({ status: 'success', message: '연동 성공! 구글 시트와 정상 연결되었습니다.' }, e.parameter.callback);
    }

    // 3. 전체 일정 목록 조회
    const schSheet = getOrCreateScheduleSheet();
    const data = schSheet.getDataRange().getValues();
    const schedules = [];

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0]) {
        let dateVal = row[3] instanceof Date ? Utilities.formatDate(row[3], Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(row[3] || '');
        let startTimeVal = row[5] instanceof Date ? Utilities.formatDate(row[5], Session.getScriptTimeZone(), 'HH:mm') : String(row[5] || '');
        let endTimeVal = row[6] instanceof Date ? Utilities.formatDate(row[6], Session.getScriptTimeZone(), 'HH:mm') : String(row[6] || '');

        schedules.push({
          id: String(row[0]),
          assignee: String(row[1] || 'husband'),
          title: String(row[2] || ''),
          date: dateVal,
          allDay: row[4] === true || String(row[4]).toUpperCase() === 'TRUE',
          startTime: startTimeVal,
          endTime: endTimeVal,
          memo: String(row[7] || '')
        });
      }
    }

    return respondJson({
      status: 'success',
      schedules: schedules
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

// 데이터 일괄 동기화
function handleDataSync(rawJson, callback) {
  const schSheet = getOrCreateScheduleSheet();
  const parsed = typeof rawJson === 'string' ? JSON.parse(rawJson) : rawJson;
  const schedules = parsed.schedules || [];
  const updatedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  // 기존 일정 영역 비우기
  const lastRow = schSheet.getLastRow();
  if (lastRow > 1) {
    schSheet.deleteRows(2, lastRow - 1);
  }

  // 새 일정 쓰기
  if (schedules.length > 0) {
    const rows = schedules.map(s => [
      s.id,
      s.assignee || 'husband',
      s.title,
      s.date,
      s.allDay ? 'TRUE' : 'FALSE',
      s.startTime || '',
      s.endTime || '',
      s.memo || '',
      updatedAt
    ]);
    schSheet.getRange(2, 1, rows.length, 9).setValues(rows);
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
