/**
 * 밴드 합주 일정 취합 — Google Apps Script 웹앱 (v2)
 *
 * 시트 3개 (스크립트가 알아서 만듭니다)
 *   설정  : A=key, B=value   → dates / hourStart / hourEnd / title
 *   응답  : A=이름, B~=날짜(YYYY-MM-DD), 값 = 가능 시각 목록 "10,11,12"
 *           "-" = 그 날 전부 불가 (응답은 했다는 뜻), 빈칸 = 미응답
 *   메모  : 응답 시트와 같은 배치, 값 = 그 날에 대한 한 줄 사유
 *
 * 통신은 JSONP(GET + callback).
 * Apps Script /exec 는 script.googleusercontent.com 으로 리다이렉트되는데
 * 브라우저 fetch 는 그 리다이렉트에서 CORS 로 응답이 멈춥니다(실측).
 * <script> 태그로 부르는 JSONP 는 그 제약을 받지 않습니다.
 *
 * 배포: 시트 → 확장 프로그램 → Apps Script → 붙여넣기 → 저장
 *      → 배포 → 배포 관리 → 연필(수정) → 버전: 새 버전 → 배포
 *      ('새 배포'를 쓰면 URL 이 바뀝니다. 반드시 기존 배포를 수정)
 */

var SHEET_ID = '1Yr8JkhPj9SzkGAKGbj699L20_svbS0hPbrnbLSc_4FM';
var S_RESP   = '응답';
var S_NOTE   = '메모';
var S_CONF   = '설정';

var DEFAULTS = { dates: '', hourStart: 9, hourEnd: 22, title: '애쉬만루트 합주' };

/* ───────────────────────── 시트 유틸 ───────────────────────── */

function ss_() { return SpreadsheetApp.openById(SHEET_ID); }

function grid_(name) {
  var ss = ss_();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1).setValue('이름');
    sh.setFrozenRows(1);
    sh.setFrozenColumns(1);
  }
  return sh;
}

function conf_() {
  var ss = ss_();
  var sh = ss.getSheetByName(S_CONF);
  if (!sh) {
    sh = ss.insertSheet(S_CONF);
    sh.getRange(1, 1, 1, 2).setValues([['key', 'value']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function confRead_() {
  var sh = conf_();
  var out = {};
  for (var k in DEFAULTS) out[k] = DEFAULTS[k];
  var last = sh.getLastRow();
  if (last >= 2) {
    var rows = sh.getRange(2, 1, last - 1, 2).getValues();
    for (var i = 0; i < rows.length; i++) {
      var k = String(rows[i][0] || '').trim();
      if (!k) continue;
      out[k] = rows[i][1];
    }
  }
  return {
    dates:     splitDates_(out.dates),
    hourStart: clampHour_(out.hourStart, DEFAULTS.hourStart),
    hourEnd:   clampHour_(out.hourEnd,   DEFAULTS.hourEnd),
    title:     String(out.title || DEFAULTS.title)
  };
}

function confWrite_(key, value) {
  var sh = conf_();
  var last = sh.getLastRow();
  if (last >= 2) {
    var keys = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < keys.length; i++) {
      if (String(keys[i][0]).trim() === key) {
        sh.getRange(i + 2, 2).setNumberFormat('@').setValue(String(value));
        return;
      }
    }
  }
  sh.getRange(Math.max(last, 1) + 1, 1, 1, 2).setNumberFormat('@')
    .setValues([[key, String(value)]]);
}

function clampHour_(v, dflt) {
  var n = parseInt(v, 10);
  if (isNaN(n) || n < 0 || n > 24) return dflt;
  return n;
}

function normDate_(v) {
  if (v === null || v === undefined || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd');
  }
  var s = String(v).trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}

function splitDates_(v) {
  var s = String(v == null ? '' : v);
  var parts = s.split(/[,\s]+/);
  var seen = {}, out = [];
  for (var i = 0; i < parts.length; i++) {
    var d = normDate_(parts[i]);
    if (d && !seen[d]) { seen[d] = 1; out.push(d); }
  }
  out.sort();
  return out;
}

/** 셀 값 → 가능 시각 배열. 구버전(O / X / 09-12) 도 읽어줍니다. */
function parseHours_(raw, cf) {
  var s = String(raw == null ? '' : raw).trim();
  if (!s) return null;                       // 미응답
  if (s === '-' || s.toUpperCase() === 'X' || s === '✕') return [];
  if (s.toUpperCase() === 'O' || s === '○') return [10, 11];

  var out = [], seen = {};
  var parts = s.split(/[,\s]+/);
  for (var i = 0; i < parts.length; i++) {
    var p = parts[i].trim();
    if (!p) continue;
    var m = p.match(/^(\d{1,2})\s*-\s*(\d{1,2})$/);   // 09-12 = 9,10,11
    if (m) {
      var a = parseInt(m[1], 10), b = parseInt(m[2], 10);
      for (var h = a; h < b; h++) if (!seen[h]) { seen[h] = 1; out.push(h); }
      continue;
    }
    var n = parseInt(p, 10);
    if (!isNaN(n) && !seen[n]) { seen[n] = 1; out.push(n); }
  }
  out.sort(function (x, y) { return x - y; });
  return out;
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

function ensureCol_(sh, map, date) {
  if (map[date]) return map[date];
  var col = Math.max(sh.getLastColumn(), 1) + 1;
  sh.getRange(1, col).setNumberFormat('@').setValue(date);
  map[date] = col;
  return col;
}

function findRow_(sh, name) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var names = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var r = 0; r < names.length; r++) {
    if (String(names[r][0]).trim() === name) return r + 2;
  }
  return 0;
}

function ensureRow_(sh, name) {
  var r = findRow_(sh, name);
  if (r) return r;
  r = Math.max(sh.getLastRow(), 1) + 1;
  sh.getRange(r, 1).setValue(name);
  return r;
}

/* ───────────────────────── 상태 읽기 ───────────────────────── */

function readGrid_(sh) {
  var lastRow = sh.getLastRow(), lastCol = Math.max(sh.getLastColumn(), 1);
  var res = { names: [], cells: {} };
  if (lastRow < 2) return res;

  var values = sh.getRange(1, 1, lastRow, lastCol).getValues();
  var head = values[0];
  for (var r = 1; r < values.length; r++) {
    var name = String(values[r][0] == null ? '' : values[r][0]).trim();
    if (!name) continue;
    res.names.push(name);
    var row = {};
    for (var c = 1; c < head.length; c++) {
      var key = normDate_(head[c]);
      if (!key) continue;
      var v = values[r][c];
      if (v === '' || v === null || v === undefined) continue;
      row[key] = String(v).trim();
    }
    res.cells[name] = row;
  }
  return res;
}

function state_() {
  var cf = confRead_();
  var resp = readGrid_(grid_(S_RESP));
  var note = readGrid_(grid_(S_NOTE));

  var hours = {};
  for (var i = 0; i < resp.names.length; i++) {
    var n = resp.names[i], row = resp.cells[n] || {}, out = {};
    for (var d in row) {
      var hh = parseHours_(row[d], cf);
      if (hh !== null) out[d] = hh;
    }
    hours[n] = out;
  }

  var notes = {};
  for (var j = 0; j < resp.names.length; j++) {
    var nm = resp.names[j];
    notes[nm] = note.cells[nm] || {};
  }

  return {
    ok: true,
    config: cf,
    members: resp.names,
    hours: hours,
    notes: notes,
    now: Utilities.formatDate(new Date(), 'Asia/Seoul', "yyyy-MM-dd'T'HH:mm:ss")
  };
}

/* ───────────────────────── 동작 ───────────────────────── */

function handle_(p) {
  var action = p.action || 'load';
  if (action === 'load') return state_();

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);

    if (action === 'save') {
      var name = String(p.name || '').trim();
      if (!name) return { ok: false, error: '이름이 필요합니다' };

      var body   = JSON.parse(p.payload || '{}');
      var pHours = body.hours || {};
      var pNotes = body.notes || {};

      var shR = grid_(S_RESP), shN = grid_(S_NOTE);
      var mapR = headerMap_(shR), mapN = headerMap_(shN);
      var rowR = ensureRow_(shR, name), rowN = ensureRow_(shN, name);

      var d;
      for (d in pHours) {
        var col = ensureCol_(shR, mapR, d);
        var arr = pHours[d];
        var val;
        if (arr === null || arr === undefined) val = '';
        else if (!arr.length) val = '-';
        else val = arr.join(',');
        shR.getRange(rowR, col).setNumberFormat('@').setValue(val);
      }
      for (d in pNotes) {
        var colN = ensureCol_(shN, mapN, d);
        shN.getRange(rowN, colN).setNumberFormat('@')
           .setValue(String(pNotes[d] || '').slice(0, 200));
      }
      shR.getRange(rowR, 2, 1, Math.max(shR.getLastColumn() - 1, 1))
         .setHorizontalAlignment('center');

    } else if (action === 'member_add') {
      var nm = String(p.name || '').trim();
      if (!nm) return { ok: false, error: '이름이 필요합니다' };
      if (findRow_(grid_(S_RESP), nm)) return { ok: false, error: '이미 있는 이름입니다' };
      ensureRow_(grid_(S_RESP), nm);
      ensureRow_(grid_(S_NOTE), nm);

    } else if (action === 'member_rename') {
      var from = String(p.from || '').trim(), to = String(p.to || '').trim();
      if (!from || !to) return { ok: false, error: '이름이 필요합니다' };
      var shR2 = grid_(S_RESP), shN2 = grid_(S_NOTE);
      if (findRow_(shR2, to)) return { ok: false, error: '이미 있는 이름입니다' };
      var r1 = findRow_(shR2, from);
      if (!r1) return { ok: false, error: '없는 멤버입니다' };
      shR2.getRange(r1, 1).setValue(to);
      var r2 = findRow_(shN2, from);
      if (r2) shN2.getRange(r2, 1).setValue(to);

    } else if (action === 'member_remove') {
      var rm = String(p.name || '').trim();
      var shR3 = grid_(S_RESP), shN3 = grid_(S_NOTE);
      var r3 = findRow_(shR3, rm);
      if (!r3) return { ok: false, error: '없는 멤버입니다' };
      shR3.deleteRow(r3);
      var r4 = findRow_(shN3, rm);
      if (r4) shN3.deleteRow(r4);

    } else if (action === 'config_set') {
      if (p.dates !== undefined)     confWrite_('dates', splitDates_(p.dates).join(','));
      if (p.hourStart !== undefined) confWrite_('hourStart', clampHour_(p.hourStart, DEFAULTS.hourStart));
      if (p.hourEnd !== undefined)   confWrite_('hourEnd',   clampHour_(p.hourEnd,   DEFAULTS.hourEnd));
      if (p.title !== undefined)     confWrite_('title', String(p.title).slice(0, 60));

    } else if (action === 'reset_answers') {
      // 날짜 후보에 대한 모든 응답·메모를 비웁니다 (멤버 목록은 유지)
      var shR4 = grid_(S_RESP), shN4 = grid_(S_NOTE);
      var lrR = shR4.getLastRow(), lcR = shR4.getLastColumn();
      if (lrR >= 2 && lcR >= 2) shR4.getRange(2, 2, lrR - 1, lcR - 1).clearContent();
      var lrN = shN4.getLastRow(), lcN = shN4.getLastColumn();
      if (lrN >= 2 && lcN >= 2) shN4.getRange(2, 2, lrN - 1, lcN - 1).clearContent();

    } else {
      return { ok: false, error: 'unknown action' };
    }

    SpreadsheetApp.flush();
    return state_();

  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

/* ───────────────────────── 엔트리 ───────────────────────── */

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
  try { out = handle_(JSON.parse(e.postData.contents)); }
  catch (err) { out = { ok: false, error: String(err) }; }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}
