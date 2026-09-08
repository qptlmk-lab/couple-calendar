/**
 * ==============================================================================
 * 💑 부부 공유 캘린더 - Google Apps Script 연동 코드
 * ==============================================================================
 * 
 * [초간단 설정 방법 (1분 소요)]
 * 1. 구글 스프레드시트(Google Sheets)를 새로 하나 만듭니다. (제목: 부부 캘린더 등)
 * 2. 상단 메뉴에서 [확장 프로그램] -> [Apps Script]를 클릭합니다.
 * 3. 기존 코드를 모두 지우고 이 파일의 전체 코드를 복사해서 붙여넣습니다.
 * 4. 상단 [저장(💾)] 버튼을 누릅니다.
 * 5. 우측 상단 파란색 [배포] -> [새 배포] 클릭
 *    - 유형 선택: [웹 앱] (톱니바퀴 아이콘)
 *    - 설명: 부부 캘린더 연동
 *    - 다음 사용자로 실행: [나 (내 계정)]
 *    - 액세스 권한이 있는 사용자: [모든 사용자 (Anyone)]  <-- ⭐ 중요!
 * 6. [배포] 버튼 클릭 후 권한 승인 진행
 * 7. 화면에 나오는 "웹 앱 URL" (https://script.google.com/macros/s/.../exec)을 복사하여
 *    부부 캘린더 앱의 [⚙️ 구글 시트 설정]에 붙여넣기만 하면 연동 끝!
 * ==============================================================================
 */

// 시트 이름
const SHEET_NAME = 'Schedules';

// 시트 초기화 및 헤더 생성
function getOrCreateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // 첫 행에 컬럼 헤더 추가
    sheet.appendRow(['id', 'title', 'date', 'allDay', 'startTime', 'endTime', 'assignee', 'category', 'completed', 'memo', 'updatedAt']);
    sheet.getRange(1, 1, 1, 11).setFontWeight('bold').setBackground('#f3e8ff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// GET 요청 처리 (전체 일정 불러오기)
function doGet(e) {
  try {
    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    const schedules = [];

    // 1행(헤더) 제외하고 2행부터 데이터 읽기
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0]) { // id가 있는 경우
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

// POST 요청 처리 (일정 추가, 수정, 삭제, 전체 동기화)
function doPost(e) {
  try {
    const sheet = getOrCreateSheet();
    let requestData;

    if (e.postData && e.postData.contents) {
      requestData = JSON.parse(e.postData.contents);
    } else if (e.parameter && e.parameter.data) {
      requestData = JSON.parse(e.parameter.data);
    } else {
      throw new Error('No data received');
    }

    const action = requestData.action || 'sync';
    const updatedAt = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    if (action === 'sync') {
      // 전체 일정 배열을 시트에 덮어쓰기 (가장 안전하고 확실한 동기화)
      const schedules = requestData.schedules || [];
      
      // 기존 데이터 삭제 (헤더 유지)
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
        message: 'Sync completed',
        count: schedules.length
      })).setMimeType(ContentService.MimeType.JSON);
    }

    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: 'Unknown action'
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
