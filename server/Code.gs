/**
 * 마음월드 데이터 서버 (Google Apps Script, 무료)
 * 만든 사람: 이경미 선생님 (elisa30@naver.com)
 *
 * 사용법
 * 1) 새 구글 시트를 만들고  [확장 프로그램] → [Apps Script] 를 연다.
 * 2) 이 파일 내용을 전부 붙여넣고 저장한다.
 * 3) 위쪽 함수 목록에서 setup 을 골라 [실행] → 권한 허용. (시트 탭 2개가 자동으로 만들어짐)
 * 4) [배포] → [새 배포] → 유형: 웹 앱 / 실행: 나 / 액세스: 모든 사용자 → 배포
 * 5) 나온 웹 앱 주소(…/exec)를 index.html 의 API_BASE 에 붙여넣는다.
 */
// 시트에 연결된 스크립트라면 비워 두세요. (다른 시트를 쓰려면 시트 ID를 적어요)
const SPREADSHEET_ID = '';
const STUDENT_SHEET = '학생정보';
const CHECKIN_SHEET = '체크인기록';

function doGet(e) {
  return handle(e);
}
function doPost(e) {
  return handle(e);
}

function handle(e) {
  try {
    const params = Object.assign({}, (e && e.parameter) || {});
    if (e && e.postData && e.postData.contents) {
      try {
        const body = JSON.parse(e.postData.contents);
        Object.assign(params, body);
      } catch (err) {}
    }
    const action = params.action;
    let result;
    switch (action) {
      case 'checkStudent':
      result = checkStudent(params);
      break;
      case 'createStudent':
      result = createStudent(params);
      break;
      case 'logCheckin':
      result = logCheckin(params);
      break;
      case 'getHistory':
      result = getHistory(params);
      break;
      default:
      result = { ok: false, error: 'unknown action: ' + action };
    }
    return jsonOutput(result);
  } catch (err) {
    return jsonOutput({ ok: false, error: String(err) });
  }
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
  .setMimeType(ContentService.MimeType.JSON);
}

const HEADERS = {
  '학생정보': ['이름', '반', '캐릭터', '첫방문', '최근방문', '레벨', '마음씨앗', '연속출석', '최근방문일'],
  '체크인기록': ['이름', '반', '일시', '감정', '장소', '활동', '보상', '메모']
};

function getSpreadsheet() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

// 탭이 없으면 자동으로 만들고 제목 줄을 넣는다.
function getSheet(name) {
  const ss = getSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow(HEADERS[name]);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEADERS[name].length).setFontWeight('bold').setBackground('#fde4ec');
    sh.getRange('B:B').setNumberFormat('@'); // 반(3-1)이 날짜로 바뀌지 않게
  }
  return sh;
}

// 처음 한 번 실행: 시트 탭 2개 준비
function setup() {
  getSheet(STUDENT_SHEET);
  getSheet(CHECKIN_SHEET);
  return '준비 완료';
}

function findStudentRow(sheet, name, cls) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const rowName = String(data[i][0] || '').trim();
    const rowCls = String(data[i][1] || '').trim();
    if (rowName === String(name).trim() && (!cls || rowCls === String(cls).trim())) {
      return { rowIndex: i + 1, row: data[i] };
    }
  }
  return null;
}

function checkStudent(params) {
  const name = params.name;
  const cls = params.cls || '';
  if (!name) return { ok: false, error: 'name required' };
  const sheet = getSheet(STUDENT_SHEET);
  const found = findStudentRow(sheet, name, cls);
  if (!found) {
    return { ok: true, exists: false };
  }
  const r = found.row;
  return {
    ok: true,
    exists: true,
    student: {
      name: r[0], cls: r[1], character: r[2],
      firstVisit: r[3], lastVisit: r[4],
      level: r[5], seeds: r[6], streak: r[7], lastVisitDate: r[8]
    }
  };
}

function createStudent(params) {
  const name = params.name;
  const cls = params.cls || '';
  const character = params.character || '';
  if (!name) return { ok: false, error: 'name required' };
  const sheet = getSheet(STUDENT_SHEET);
  const existing = findStudentRow(sheet, name, cls);
  if (existing) {
    return { ok: false, error: 'student already exists' };
  }
  const now = new Date();
  sheet.appendRow([name, cls, character, now, now, 1, 0, 1, now]);
  const newRow = sheet.getLastRow();
  sheet.getRange(newRow, 2).setNumberFormat('@').setValue(cls);
  return { ok: true };
}

function logCheckin(params) {
  const name = params.name;
  const cls = params.cls || '';
  if (!name) return { ok: false, error: 'name required' };
  const checkinSheet = getSheet(CHECKIN_SHEET);
  const now = new Date();
  checkinSheet.appendRow([
      name, cls, now,
      params.emotion || '',
      params.zone || '',
      params.activity || '',
      params.reward || 0,
      params.memo || ''
  ]);
  const newCheckinRow = checkinSheet.getLastRow();
  checkinSheet.getRange(newCheckinRow, 2).setNumberFormat('@').setValue(cls);

  const studentSheet = getSheet(STUDENT_SHEET);
  const found = findStudentRow(studentSheet, name, cls);
  if (found) {
    const rowIndex = found.rowIndex;
    const prevSeeds = Number(found.row[6]) || 0;
    const prevStreak = Number(found.row[7]) || 0;
    const prevLastDate = found.row[8] ? new Date(found.row[8]) : null;

    const rewardNum = Number(params.reward) || 0;
    const newSeeds = prevSeeds + rewardNum;
    const newLevel = Math.floor(newSeeds / 100) + 1;

    let newStreak = prevStreak;
    const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
    if (prevLastDate) {
      const prevDay = new Date(prevLastDate.getFullYear(), prevLastDate.getMonth(), prevLastDate.getDate());
      const diffDays = Math.round((today - prevDay) / (1000 * 60 * 60 * 24));
      if (diffDays === 1) newStreak = prevStreak + 1;
      else if (diffDays > 1) newStreak = 1;
    } else {
      newStreak = 1;
    }

    studentSheet.getRange(rowIndex, 5).setValue(now);
    studentSheet.getRange(rowIndex, 6).setValue(newLevel);
    studentSheet.getRange(rowIndex, 7).setValue(newSeeds);
    studentSheet.getRange(rowIndex, 8).setValue(newStreak);
    studentSheet.getRange(rowIndex, 9).setValue(now);

    return { ok: true, seeds: newSeeds, level: newLevel, streak: newStreak };
  }
  return { ok: true, note: 'student not found, checkin logged without update' };
}

function getHistory(params) {
  const name = params.name;
  const cls = params.cls || '';
  if (!name) return { ok: false, error: 'name required' };
  const sheet = getSheet(CHECKIN_SHEET);
  const data = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(name).trim() && (!cls || String(data[i][1]).trim() === String(cls).trim())) {
      rows.push({
          datetime: data[i][2], emotion: data[i][3], zone: data[i][4],
          activity: data[i][5], reward: data[i][6], memo: data[i][7]
      });
    }
  }
  return { ok: true, history: rows };
}
