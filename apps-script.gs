/**
 * 애쉬만루트 합주 일정 취합 — Google Apps Script 웹앱
 *
 * 시트 구조 (스크립트가 알아서 만듭니다)
 *   A열 = 이름(= 멤버 목록), B열부터 = 날짜(YYYY-MM-DD)
 *   값 = "O" | "X" | "09-12, 14-16" | 빈칸(미응답)
 *
 * 통신은 JSONP(GET + callback)로 합니다.
 * Apps Script /exec 는 script.googleusercontent.com 으로 리다이렉트되는데
 * 브라우저 fetch 로는 이 리다이렉트에서 CORS 때문에 응답이 멈춥니다.
 * <script> 태그로 부르는 JSONP 는 그 제약을 받지 않습니다.
 *
 * 배포: 확장 프로그램 → Apps Script → 붙여넣기 → 배포 → 배포 관리
 *      → 연필(수정) → 버전: 새 버전 → 배포
 *      (URL 을 유지하려면 '새 배포'가 아니라 기존 배포를 수정해야 합니다)
 */

var SHEET_ID   = '1Yr8JkhPj9SzkGAKGbj699L20_svbS0hPbrnbLSc_4FM';
var SHEET_NAME = '응답';
var SEED       = ['멤버 1', '멤버 2'];

function sheet_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1).setValue('이름');
    sh.setFrozenRows(1);
    sh.setFrozenColumns(1);
    for (var i = 0; i < SEED.length; i++) sh.getRange(2 + i, 1).setValue(SEED[i]);
  }
  return sh;
}

function normDate_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd');
  }
  var s = String(v).trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}

function headerMap_(sh) {
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var head = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  var map = {};
  for (var i = 1; i < head.length; i++) {
    var k = normDate_(head[i]);
    if (k) map[k] = i + 1;
  }
  return map;
}

function findRow_(sh, name) {
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;
  var names = sh.getRange(2, 1, lastRow - 1, 1).getValues();
  for (var r = 0; r < names.length; r++) {
    if (String(names[r][0]).trim() === name) return r + 2;
  }
  return 0;
}

/** 멤버 목록 + 전체 응답 */
function state_() {
  var sh = sheet_();
  var lastRow = sh.getLastRow(), lastCol = Math.max(sh.getLastColumn(), 1);
  var members = [], responses = {};
  if (lastRow < 2) return { ok: true, members: members, responses: responses };

  var values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  var head = values[0];
  for (var r = 1; r < values.length; r++) {
    var name = String(values[r][0] || '').trim();
    if (!name) continue;
    members.push(name);
    var row = {};
    for (var c = 1; c < head.length; c++) {
      var key = normDate_(head[c]);
      if (!key) continue;
      var val = String(values[r][c] == null ? '' : values[r][c]).trim();
      if (val) row[key] = val;
    }
    responses[name] = row;
  }
  return { ok: true, members: members, responses: responses };
}

/** 모든 동작의 실제 처리부 */
function handle_(p) {
  var action = p.action || 'load';
  if (action === 'load') return state_();

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var sh = sheet_();

    if (action === 'save') {
      var name = String(p.name || '').trim();
      if (!name) return { ok: false, error: 'name required' };
      var answers = JSON.parse(p.payload || '{}');

      var map = headerMap_(sh);
      var dates = Object.keys(answers).sort();
      for (var i = 0; i < dates.length; i++) {
        if (!map[dates[i]]) {
          var col = Math.max(sh.getLastColumn(), 1) + 1;
          sh.getRange(1, col).setNumberFormat('@').setValue(dates[i]);
          map[dates[i]] = col;
        }
      }
      var row = findRow_(sh, name);
      if (!row) { row = Math.max(sh.getLastRow(), 1) + 1; sh.getRange(row, 1).setValue(name); }
      for (var d = 0; d < dates.length; d++) {
        sh.getRange(row, map[dates[d]]).setValue(answers[dates[d]] || '');
      }
      sh.getRange(row, 2, 1, Math.max(sh.getLastColumn() - 1, 1)).setHorizontalAlignment('center');

    } else if (action === 'member_add') {
      var nm = String(p.name || '').trim();
      if (!nm) return { ok: false, error: 'name required' };
      if (findRow_(sh, nm)) return { ok: false, error: '이미 있는 이름입니다' };
      sh.getRange(Math.max(sh.getLastRow(), 1) + 1, 1).setValue(nm);

    } else if (action === 'member_rename') {
      var from = String(p.from || '').trim(), to = String(p.to || '').trim();
      if (!from || !to) return { ok: false, error: 'from/to required' };
      if (findRow_(sh, to)) return { ok: false, error: '이미 있는 이름입니다' };
      var r1 = findRow_(sh, from);
      if (!r1) return { ok: false, error: '없는 멤버입니다' };
      sh.getRange(r1, 1).setValue(to);

    } else if (action === 'member_remove') {
      var rm = String(p.name || '').trim();
      var r2 = findRow_(sh, rm);
      if (!r2) return { ok: false, error: '없는 멤버입니다' };
      sh.deleteRow(r2);

    } else {
      return { ok: false, error: 'unknown action' };
    }

    SpreadsheetApp.flush();
    return state_();

  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  var out;
  try { out = handle_(p); }
  catch (err) { out = { ok: false, error: String(err) }; }

  var body = JSON.stringify(out);
  if (p.callback) {
    return ContentService
      .createTextOutput(p.callback + '(' + body + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body)
    .setMimeType(ContentService.MimeType.JSON);
}

/** 서버 대 서버용 (curl / PowerShell). 브라우저는 doGet + JSONP 를 씁니다. */
function doPost(e) {
  var out;
  try {
    var body = JSON.parse(e.postData.contents);
    if (body.answers && !body.payload) body.payload = JSON.stringify(body.answers);
    out = handle_(body);
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}
