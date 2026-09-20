/**
 * 밴드매니저 — Google Apps Script 웹앱 (v3)
 *
 * 스프레드시트는 "데이터 저장소"일 뿐입니다. 모든 조작은 웹페이지에서 합니다.
 * 탭 8개 (스크립트가 알아서 만듭니다). 모든 행은 첫 컬럼에 밴드를 답니다.
 *
 *   밴드   id | 이름
 *   멤버   밴드 | 이름
 *   설정   밴드 | key | value           (dates / hourStart / hourEnd / title / windows / pin)
 *            windows = JSON {"2026-09-20":[9,12], ...}  날짜별 시간대 (없으면 hourStart~hourEnd)
 *            tables  = JSON [{"id":"t1","name":"시간표1","dates":[...]}, ...]  여러 시간표 (v13)
 *                      dates 는 모든 시간표 날짜의 합집합으로 같이 저장된다 (응답은 날짜 단위라 시간표끼리 공유)
 *   응답   밴드 | 이름 | 날짜 | 시간     ("10,11,12" 또는 "-" = 그 날 불가)
 *   메모   밴드 | 이름 | 날짜 | 메모
 *   곡     id | 밴드 | 상태 | 제목 | 아티스트 | 키 | 선곡자 | 링크 | 파트 | 튜닝 | 메모 | 추가일
 *            상태 = 후보 | 연습중 | 완료 | 보류
 *            링크 = 줄바꿈으로 여러 개. 각 줄 "라벨|URL" 또는 "URL"
 *   이력   id | 밴드 | 날짜 | 시작 | 종료 | 합주실 | 룸 | 상태 | 불참 | 곡 | 메모
 *            상태 = 예정 | 완료 | 취소 / 불참·곡 = 쉼표로 구분
 *   투표   밴드 | 곡 | 이름 | 값        (값 = 1 좋아요 / -1 별로 / 0 보류)
 *
 * 헤더가 다른 옛 탭이 있으면 "<이름>_구버전"으로 이름만 바꿔 보존하고 새로 만듭니다.
 *
 * 통신은 JSONP(GET + callback). fetch 는 Apps Script 리다이렉트에서 CORS 로 멈춥니다.
 *
 * 배포: 코드 교체 → Ctrl+S → 배포 → 배포 관리 → 연필(수정) → 버전: 새 버전 → 배포
 *      ('새 배포'를 쓰면 URL 이 바뀝니다. 반드시 기존 배포를 수정)
 */

var SHEET_ID = '1Yr8JkhPj9SzkGAKGbj699L20_svbS0hPbrnbLSc_4FM';

var TABS = {
  band:   { name: '밴드', head: ['id', '이름'] },
  member: { name: '멤버', head: ['밴드', '이름'] },
  conf:   { name: '설정', head: ['밴드', 'key', 'value'] },
  resp:   { name: '응답', head: ['밴드', '이름', '날짜', '시간'] },
  note:   { name: '메모', head: ['밴드', '이름', '날짜', '메모'] },
  song:   { name: '곡',   head: ['id','밴드','상태','제목','아티스트','키','선곡자','링크','파트','튜닝','메모','추가일'] },
  hist:   { name: '이력', head: ['id','밴드','날짜','시작','종료','합주실','룸','상태','불참','곡','메모'] },
  vote:   { name: '투표', head: ['밴드', '곡', '이름', '값'] }
};

var CONF_DEFAULT = { dates: '', hourStart: 9, hourEnd: 22, title: '' };

/* ═══════════════ 시트 유틸 ═══════════════ */

/* 스프레드시트 왕복이 느립니다(한 번에 100~300ms). 그래서 한 번의 실행 안에서
 * ─ Spreadsheet 객체(_SS)
 * ─ 탭 이름 → Sheet 객체 맵(_SHEETS, getSheets() 한 번으로)
 * ─ 탭별 Sheet(_TAB) 와 읽어온 행들(_ROWS)
 * 을 전부 캐시합니다. 탭 하나당 읽기는 getDataRange() 한 번뿐이고,
 * put_ 로 쓴 내용은 캐시에 그대로 반영해 두어 뒤따르는 state_() 가 다시 읽지 않습니다. */
var _SS = null, _SHEETS = null, _TAB = {}, _ROWS = {};

/* 스프레드시트 왕복 하나가 수백 ms 라, 한 요청에 8개 탭을 읽으면 6~14초가 걸립니다.
 * 그래서 탭별 행들을 ScriptCache 에도 넣어 둡니다.
 * ─ 읽기(load)는 캐시가 있으면 스프레드시트를 아예 열지 않습니다
 * ─ 쓰기는 시트에 쓰고 캐시도 같은 내용으로 갱신합니다
 * ─ 시트를 손으로 고쳤다면 `fresh=1` 로 캐시를 무시하고 다시 읽습니다 (페이지의 '새로고침') */
var CACHE_TTL = 21600;       // 6시간 (ScriptCache 최대값)
/* 캐시 적중분은 수명만 다시 늘려준다(슬라이딩 만료). 누가 6시간 안에 한 번씩만 써도
 * 캐시가 살아 있으니, 만료된 캐시를 만난 사람이 8개 탭을 전부 읽는 일(10초 이상)이 없다.
 * 호출 끝에 putAll 한 번으로 모아서 쓴다. */
var _TOUCH = {};
function touchFlush_() {
  var n = 0; for (var k in _TOUCH) n++;
  if (!n) return;
  try { cache_().putAll(_TOUCH, CACHE_TTL); } catch (e) {}
  _TOUCH = {};
}
var _FRESH = false;
function cache_() { return CacheService.getScriptCache(); }
function ckey_(key) { return 'bm3_' + key; }
function keysAll_() { var a = []; for (var k in TABS) a.push(ckey_(k)); return a; }

/** 15분마다 깨워서 캐시를 미리 채워둔다.
 *  이게 없으면 한동안 아무도 안 쓴 뒤 첫 사용자가 15초를 기다린다
 *  (컨테이너 콜드 스타트 + 캐시 만료로 8개 탭을 전부 다시 읽기).
 *  트리거: Apps Script 편집기 왼쪽 ⏰ 트리거 → 트리거 추가 →
 *         함수 warm / 시간 기반 / 분 단위 타이머 / 15분마다 */
function warm() {
  _SS = null; _SHEETS = null; _TAB = {}; _ROWS = {};
  _FRESH = true;                       // 시트에서 다시 읽어 캐시를 새로 채운다
  for (var k in TABS) open_(k);
  _FRESH = false;
  return 'warm ok';
}

function ss_() {
  if (!_SS) _SS = SpreadsheetApp.openById(SHEET_ID);
  return _SS;
}

function sheetByName_(name) {
  if (!_SHEETS) {
    _SHEETS = {};
    var all = ss_().getSheets();
    for (var i = 0; i < all.length; i++) _SHEETS[all[i].getName()] = all[i];
  }
  return _SHEETS[name] || null;
}

/** 헤더가 맞는 탭을 보장하고, 그 탭 전체를 한 번에 읽어 캐시한다. 옛 탭은 이름만 바꿔 보존. */
function open_(key) {
  if (_ROWS[key]) return _ROWS[key];

  if (!_FRESH) {                                   // 캐시 적중이면 시트를 열지 않는다
    var hit = cache_().get(ckey_(key));
    if (hit) {
      try {
        _ROWS[key] = JSON.parse(hit);
        _TOUCH[ckey_(key)] = hit;                  // 쓴 지 6시간이 지나 만료되지 않게 수명을 늘린다
        return _ROWS[key];
      } catch (e) {}
    }
  }

  var spec = TABS[key], ss = ss_(), sh = sheetByName_(spec.name), vals = null;

  if (sh) {
    if (!sh.getLastRow()) {                        // 비어 있는 탭 → 헤더만 써 넣고 그대로 쓴다
      sh.getRange(1, 1, 1, spec.head.length).setValues([spec.head]).setFontWeight('bold');
      sh.setFrozenRows(1);
      vals = [spec.head];
    } else {
      vals = sh.getDataRange().getValues();
      var head = vals[0] || [];
      var same = true;
      for (var i = 0; i < spec.head.length; i++) {
        if (String(head[i] || '').trim() !== spec.head[i]) { same = false; break; }
      }
      if (!same) {                                 // 구버전 탭 → 이름만 바꿔 보존
        var bak = spec.name + '_구버전';
        if (sheetByName_(bak)) bak += '_' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'MMddHHmm');
        sh.setName(bak);
        _SHEETS[bak] = sh;
        delete _SHEETS[spec.name];
        sh = null; vals = null;
      }
    }
  }

  if (!sh) {
    try { sh = ss.insertSheet(spec.name); }
    catch (err) { sh = ss.getSheetByName(spec.name); if (!sh) throw err; }
    sh.getRange(1, 1, 1, spec.head.length).setValues([spec.head]).setFontWeight('bold');
    sh.setFrozenRows(1);
    _SHEETS[spec.name] = sh;
    vals = [spec.head];
  }

  _TAB[key] = sh;

  var out = [];
  for (var r = 1; r < vals.length; r++) {
    var o = {}, empty = true;
    for (var c = 0; c < spec.head.length; c++) {
      var v = vals[r][c];
      if (Object.prototype.toString.call(v) === '[object Date]') {
        v = Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd');
      }
      v = (v === null || v === undefined) ? '' : String(v).trim();
      if (v) empty = false;
      o[spec.head[c]] = v;
    }
    if (!empty) out.push(o);
  }
  _ROWS[key] = out;
  cache_().put(ckey_(key), JSON.stringify(out), CACHE_TTL);
  return out;
}

/** Sheet 객체가 필요할 때만 시트를 연다 (쓰기 경로 전용). 값은 읽지 않는다. */
function sheetOf_(key) {
  if (_TAB[key]) return _TAB[key];
  var sh = sheetByName_(TABS[key].name);
  if (sh) { _TAB[key] = sh; return sh; }

  var keep = _ROWS[key];                           // 탭이 없으면 open_ 이 만들게 한다
  delete _ROWS[key];
  var save = _FRESH; _FRESH = true;
  try { open_(key); } finally { _FRESH = save; }
  if (keep) _ROWS[key] = keep;
  return _TAB[key];
}

function tab_(key) { return sheetOf_(key); }

/** 탭 전체를 객체 배열로 (실행 중엔 캐시) */
function rows_(key) { return open_(key); }

/** 탭 전체 덮어쓰기. 캐시도 같이 갱신해 뒤따르는 읽기가 시트를 안 건드리게 한다. */
function put_(key, list) {
  var spec = TABS[key];
  var sh = sheetOf_(key);

  var out = [];
  for (var r = 0; r < list.length; r++) {
    var row = [];
    for (var c = 0; c < spec.head.length; c++) {
      var v = list[r][spec.head[c]];
      row.push(v === null || v === undefined ? '' : String(v));
    }
    out.push(row);
  }

  var lastRow = sh.getLastRow();
  if (out.length) sh.getRange(2, 1, out.length, spec.head.length).setNumberFormat('@').setValues(out);
  if (lastRow > out.length + 1) {                  // 줄어든 만큼만 지운다
    sh.getRange(out.length + 2, 1, lastRow - out.length - 1, spec.head.length).clearContent();
  }

  _ROWS[key] = list;
  cache_().put(ckey_(key), JSON.stringify(list), CACHE_TTL);
  /* 🚨 2026-09-20 사고: 읽을 때 _TOUCH 에 잡아둔 '옛 값'이 호출 끝 touchFlush_() 에서 캐시에 다시 써져
     방금 쓴 내용을 덮었다 → 다음 쓰기가 옛 목록을 읽어 시트까지 되돌림. 쓴 탭은 touch 목록에서 뺀다. */
  delete _TOUCH[ckey_(key)];
}

function uid_(p) {
  return p + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyMMddHHmmss') +
         Math.floor(Math.random() * 1000);
}

function today_() { return Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd'); }

function normDate_(s) {
  s = String(s || '').trim();
  return /^\d{4}-\d{2}-\d{2}$/.test(s) ? s : '';
}

function splitList_(s) {
  var parts = String(s || '').split(/[,\n]/), out = [];
  for (var i = 0; i < parts.length; i++) {
    var p = parts[i].trim();
    if (p) out.push(p);
  }
  return out;
}

function clampHour_(v, d) {
  var n = parseInt(v, 10);
  return (isNaN(n) || n < 0 || n > 24) ? d : n;
}

/** 날짜별 시간대 JSON 을 검증해서 {날짜: [시작, 끝]} 만 남긴다. 이상한 건 버린다. */
function parseWindows_(raw) {
  var out = {}, o;
  try { o = typeof raw === 'string' ? JSON.parse(raw || '{}') : (raw || {}); } catch (e) { return out; }
  if (!o || typeof o !== 'object') return out;
  var n = 0;
  for (var k in o) {
    var d = normDate_(k); if (!d) continue;
    var w = o[k]; if (!w || w.length !== 2) continue;
    var a = clampHour_(w[0], -1), b = clampHour_(w[1], -1);
    if (a < 0 || b < 0 || b <= a) continue;
    out[d] = [a, b];
    if (++n >= 120) break;
  }
  return out;
}

/** "10,11,12" → [10,11,12] / "-" → [] / 빈칸 → null(미응답) */
/** 시간표 목록 검증: 최대 20개, 이름 30자, 날짜 정규화·중복 제거·정렬 */
function parseTables_(raw) {
  var arr;
  try { arr = typeof raw === 'string' ? JSON.parse(raw) : raw; } catch (e) { return []; }
  if (!arr || !arr.length) return [];
  var out = [], ids = {};
  for (var i = 0; i < arr.length && out.length < 20; i++) {
    var t = arr[i] || {};
    var id = String(t.id || '').replace(/[^A-Za-z0-9_-]/g, '').slice(0, 24) || ('t' + (i + 1));
    if (ids[id]) continue; ids[id] = 1;
    var ds = t.dates || [], keep = [], seen = {};
    for (var j = 0; j < ds.length; j++) {
      var nd = normDate_(ds[j]);
      if (nd && !seen[nd]) { seen[nd] = 1; keep.push(nd); }
    }
    keep.sort();
    out.push({ id: id, name: String(t.name || ('시간표' + (i + 1))).slice(0, 30), dates: keep });
  }
  return out;
}

function parseHours_(raw) {
  var s = String(raw == null ? '' : raw).trim();
  if (!s) return null;
  if (s === '-') return [];
  var out = [], seen = {}, parts = s.split(/[,\s]+/);
  for (var i = 0; i < parts.length; i++) {
    var n = parseInt(parts[i], 10);
    if (!isNaN(n) && !seen[n]) { seen[n] = 1; out.push(n); }
  }
  out.sort(function (a, b) { return a - b; });
  return out;
}

/* ═══════════════ 상태 ═══════════════ */

function bands_() {
  var list = rows_('band');
  if (!list.length) {
    list = [{ id: 'band1', 이름: '우리 밴드' }];
    put_('band', list);
  }
  return list;
}

/* ─── 밴드 PIN ───────────────────────────────────────────────
 * 설정 탭의 `pin` 키에 저장한다. PIN 이 걸린 밴드는
 * 맞는 PIN 을 보내지 않으면 **데이터를 아예 내려주지 않는다**(state_ 에서 제외).
 * 쓰기도 마찬가지로 막는다. 클라이언트는 `pins` 에 {밴드id: pin} 을 담아 보낸다. */
function pinsOf_(p) {
  try { return JSON.parse(p.pins || '{}') || {}; } catch (e) { return {}; }
}

function bandPin_(bandId) {
  var cs = rows_('conf');
  for (var i = 0; i < cs.length; i++) {
    if (cs[i]['밴드'] === bandId && cs[i]['key'] === 'pin') return String(cs[i]['value'] || '');
  }
  return '';
}

function allowed_(bandId, pins) {
  var need = bandPin_(bandId);
  if (!need) return true;
  return String(pins[bandId] == null ? '' : pins[bandId]) === need;
}

function state_(pins) {
  pins = pins || {};
  var bl = bands_();
  var open = {};                                   // 열려 있는(= 데이터를 내려줄) 밴드
  var out = {
    ok: true,
    bands: [], members: {}, config: {}, hours: {}, notes: {},
    songs: [], hist: [], votes: {},
    now: Utilities.formatDate(new Date(), 'Asia/Seoul', "yyyy-MM-dd'T'HH:mm:ss")
  };

  for (var i = 0; i < bl.length; i++) {
    var b = bl[i].id;
    var has = !!bandPin_(b);
    var ok  = allowed_(b, pins);
    if (ok) open[b] = true;
    out.bands.push({ id: b, name: bl[i]['이름'], hasPin: has, locked: !ok });
    out.members[b] = [];
    out.config[b] = { dates: [], hourStart: CONF_DEFAULT.hourStart,
                      hourEnd: CONF_DEFAULT.hourEnd, title: bl[i]['이름'], windows: {} };
    out.hours[b] = {};
    out.notes[b] = {};
  }

  var ms = rows_('member'), i2;
  for (i2 = 0; i2 < ms.length; i2++) {
    var mb = ms[i2]['밴드'];
    if (open[mb] && out.members[mb]) out.members[mb].push(ms[i2]['이름']);
  }

  var cs = rows_('conf');
  for (i2 = 0; i2 < cs.length; i2++) {
    var cb = cs[i2]['밴드'], k = cs[i2]['key'], v = cs[i2]['value'];
    if (!out.config[cb] || !open[cb]) continue;    // pin 값은 어차피 아래에서 안 읽는다
    if (k === 'dates') {
      var ds = splitList_(v), keep = [], seen = {};
      for (var d = 0; d < ds.length; d++) {
        var nd = normDate_(ds[d]);
        if (nd && !seen[nd]) { seen[nd] = 1; keep.push(nd); }
      }
      keep.sort();
      out.config[cb].dates = keep;
    } else if (k === 'hourStart') out.config[cb].hourStart = clampHour_(v, CONF_DEFAULT.hourStart);
    else if (k === 'hourEnd')     out.config[cb].hourEnd   = clampHour_(v, CONF_DEFAULT.hourEnd);
    else if (k === 'title')       out.config[cb].title     = String(v || '');
    else if (k === 'windows')     out.config[cb].windows   = parseWindows_(v);
    else if (k === 'tables')      out.config[cb].tables    = parseTables_(v);
  }

  var rs = rows_('resp');
  for (i2 = 0; i2 < rs.length; i2++) {
    var rb = rs[i2]['밴드'], rn = rs[i2]['이름'], rd = normDate_(rs[i2]['날짜']);
    if (!out.hours[rb] || !open[rb] || !rn || !rd) continue;
    var hh = parseHours_(rs[i2]['시간']);
    if (hh === null) continue;
    if (!out.hours[rb][rn]) out.hours[rb][rn] = {};
    out.hours[rb][rn][rd] = hh;
  }

  var ns = rows_('note');
  for (i2 = 0; i2 < ns.length; i2++) {
    var nb = ns[i2]['밴드'], nn = ns[i2]['이름'], ndd = normDate_(ns[i2]['날짜']);
    if (!out.notes[nb] || !open[nb] || !nn || !ndd || !ns[i2]['메모']) continue;
    if (!out.notes[nb][nn]) out.notes[nb][nn] = {};
    out.notes[nb][nn][ndd] = ns[i2]['메모'];
  }

  var ss = rows_('song');
  for (i2 = 0; i2 < ss.length; i2++) {
    var s = ss[i2];
    if (!open[s['밴드']]) continue;
    out.songs.push({
      id: s.id, band: s['밴드'], status: s['상태'] || '후보',
      title: s['제목'], artist: s['아티스트'], key: s['키'], picker: s['선곡자'],
      links: s['링크'], parts: s['파트'], tuning: s['튜닝'],
      memo: s['메모'], added: s['추가일']
    });
  }

  var hs = rows_('hist');
  for (i2 = 0; i2 < hs.length; i2++) {
    var h = hs[i2];
    if (!open[h['밴드']]) continue;
    out.hist.push({
      id: h.id, band: h['밴드'], date: normDate_(h['날짜']),
      from: h['시작'], to: h['종료'], place: h['합주실'], room: h['룸'],
      status: h['상태'] || '예정',
      absent: splitList_(h['불참']), songs: splitList_(h['곡']), memo: h['메모']
    });
  }
  out.hist.sort(function (a, b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : 0; });

  var vs = rows_('vote');
  for (i2 = 0; i2 < vs.length; i2++) {
    var vsid = vs[i2]['곡'], vn = vs[i2]['이름'];
    if (!vsid || !vn || !open[vs[i2]['밴드']]) continue;
    if (!out.votes[vsid]) out.votes[vsid] = {};
    out.votes[vsid][vn] = parseInt(vs[i2]['값'], 10) || 0;
  }

  return out;
}

/* ═══════════════ 동작 ═══════════════ */

function handle_(p) {
  var action = p.action || 'load';
  var pins = pinsOf_(p);
  if (p.fresh) { _FRESH = true; cache_().removeAll(keysAll_()); }   // 시트를 손으로 고쳤을 때
  if (action === 'load') { var st = state_(pins); touchFlush_(); return st; }

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    // 밴드가 걸린 동작은 그 밴드의 PIN 을 알아야 한다 (band_add 는 대상 밴드가 없다)
    if (action !== 'band_add' && !allowed_(String(p.band || ''), pins)) {
      return { ok: false, error: 'PIN이 맞지 않습니다' };
    }
    var r = act_(action, p);
    if (r && r.ok === false) return r;
    // PIN 을 방금 바꿨으면 그 값으로 응답해야 화면이 잠기지 않는다
    if (p.setpin !== undefined) pins[String(p.band || '')] = String(p.setpin).trim();
    SpreadsheetApp.flush();
    var out = state_(pins); touchFlush_(); return out;
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

function act_(action, p) {
  var band = String(p.band || '').trim();
  var i, list;

  /* ── 밴드 ── */
  if (action === 'band_add') {
    var bn = String(p.name || '').trim();
    if (!bn) return { ok: false, error: '밴드 이름이 필요합니다' };
    list = bands_();
    for (i = 0; i < list.length; i++) if (list[i]['이름'] === bn) return { ok: false, error: '이미 있는 밴드입니다' };
    list.push({ id: uid_('b'), 이름: bn });
    put_('band', list);
    return;
  }
  if (action === 'band_rename') {
    list = bands_();
    for (i = 0; i < list.length; i++) if (list[i].id === band) list[i]['이름'] = String(p.name || '').trim();
    put_('band', list);
    return;
  }
  if (action === 'band_remove') {
    list = bands_();
    if (list.length <= 1) return { ok: false, error: '마지막 밴드는 지울 수 없습니다' };
    put_('band', list.filter(function (x) { return x.id !== band; }));
    ['member','conf','resp','note','song','hist','vote'].forEach(function (k) {
      var f = TABS[k].head[0] === 'id' ? '밴드' : '밴드';
      put_(k, rows_(k).filter(function (x) { return x[f] !== band; }));
    });
    return;
  }

  /* ── 멤버 ── */
  if (action === 'member_add') {
    var mn = String(p.name || '').trim();
    if (!mn) return { ok: false, error: '이름이 필요합니다' };
    list = rows_('member');
    for (i = 0; i < list.length; i++)
      if (list[i]['밴드'] === band && list[i]['이름'] === mn) return { ok: false, error: '이미 있는 이름입니다' };
    list.push({ '밴드': band, '이름': mn });
    put_('member', list);
    return;
  }
  if (action === 'member_rename') {
    var from = String(p.from || '').trim(), to = String(p.to || '').trim();
    if (!from || !to) return { ok: false, error: '이름이 필요합니다' };
    list = rows_('member');
    for (i = 0; i < list.length; i++)
      if (list[i]['밴드'] === band && list[i]['이름'] === to) return { ok: false, error: '이미 있는 이름입니다' };
    var found = false;
    for (i = 0; i < list.length; i++)
      if (list[i]['밴드'] === band && list[i]['이름'] === from) { list[i]['이름'] = to; found = true; }
    if (!found) return { ok: false, error: '없는 멤버입니다' };
    put_('member', list);
    ['resp','note','vote'].forEach(function (k) {
      var l = rows_(k);
      for (var j = 0; j < l.length; j++)
        if (l[j]['밴드'] === band && l[j]['이름'] === from) l[j]['이름'] = to;
      put_(k, l);
    });
    return;
  }
  if (action === 'member_remove') {
    var rm = String(p.name || '').trim();
    put_('member', rows_('member').filter(function (x) {
      return !(x['밴드'] === band && x['이름'] === rm);
    }));
    ['resp','note','vote'].forEach(function (k) {
      put_(k, rows_(k).filter(function (x) { return !(x['밴드'] === band && x['이름'] === rm); }));
    });
    return;
  }

  /* ── 설정 ── */
  if (action === 'config_set') {
    list = rows_('conf');
    var set = function (k, v) {
      for (var j = 0; j < list.length; j++)
        if (list[j]['밴드'] === band && list[j]['key'] === k) { list[j]['value'] = String(v); return; }
      list.push({ '밴드': band, 'key': k, 'value': String(v) });
    };
    if (p.dates !== undefined) {
      var ds = splitList_(p.dates), keep = [], seen = {};
      for (i = 0; i < ds.length; i++) {
        var nd = normDate_(ds[i]);
        if (nd && !seen[nd]) { seen[nd] = 1; keep.push(nd); }
      }
      keep.sort(); set('dates', keep.join(','));
    }
    if (p.hourStart !== undefined) set('hourStart', clampHour_(p.hourStart, CONF_DEFAULT.hourStart));
    if (p.hourEnd   !== undefined) set('hourEnd',   clampHour_(p.hourEnd,   CONF_DEFAULT.hourEnd));
    if (p.title     !== undefined) set('title', String(p.title).slice(0, 60));
    if (p.windows   !== undefined) set('windows', JSON.stringify(parseWindows_(p.windows)));
    if (p.tables    !== undefined) {                 // 여러 시간표 — dates 는 합집합으로 맞춘다
      var tl = parseTables_(p.tables), un = {}, ua = [];
      for (i = 0; i < tl.length; i++) for (var q = 0; q < tl[i].dates.length; q++) un[tl[i].dates[q]] = 1;
      for (var kk in un) ua.push(kk);
      ua.sort();
      set('tables', JSON.stringify(tl));
      set('dates', ua.join(','));
    }
    if (p.setpin    !== undefined) set('pin',   String(p.setpin).trim().slice(0, 20));
    put_('conf', list);
    return;
  }

  /* ── 응답 + 메모 ── */
  if (action === 'save') {
    var name = String(p.name || '').trim();
    if (!name) return { ok: false, error: '이름이 필요합니다' };
    var body = JSON.parse(p.payload || '{}');
    var ph = body.hours || {}, pn = body.notes || {};

    var rl = rows_('resp').filter(function (x) {
      return !(x['밴드'] === band && x['이름'] === name && ph[x['날짜']] !== undefined);
    });
    for (var d1 in ph) {
      if (!normDate_(d1)) continue;
      var arr = ph[d1];
      if (arr === null) continue;          // null = 아직 안 정함 → 줄을 지우기만 한다 (v12)
      rl.push({ '밴드': band, '이름': name, '날짜': d1,
                '시간': (!arr || !arr.length) ? '-' : arr.join(',') });
    }
    put_('resp', rl);

    var nl = rows_('note').filter(function (x) {
      return !(x['밴드'] === band && x['이름'] === name && pn[x['날짜']] !== undefined);
    });
    for (var d2 in pn) {
      if (!normDate_(d2) || !pn[d2]) continue;
      nl.push({ '밴드': band, '이름': name, '날짜': d2, '메모': String(pn[d2]).slice(0, 200) });
    }
    put_('note', nl);
    return;
  }

  if (action === 'reset_answers') {
    put_('resp', rows_('resp').filter(function (x) { return x['밴드'] !== band; }));
    put_('note', rows_('note').filter(function (x) { return x['밴드'] !== band; }));
    return;
  }

  /* ── 곡 ── */
  if (action === 'song_save') {
    var s = JSON.parse(p.payload || '{}');
    if (!String(s.title || '').trim()) return { ok: false, error: '곡 제목이 필요합니다' };
    list = rows_('song');
    var row = {
      id: s.id || uid_('s'), '밴드': band, '상태': s.status || '후보',
      '제목': s.title, '아티스트': s.artist || '', '키': s.key || '',
      '선곡자': s.picker || '', '링크': s.links || '', '파트': s.parts || '',
      '튜닝': s.tuning || '', '메모': s.memo || '', '추가일': s.added || today_()
    };
    var hit = false;
    for (i = 0; i < list.length; i++) if (list[i].id === row.id) { list[i] = row; hit = true; }
    if (!hit) list.push(row);
    put_('song', list);
    return;
  }
  if (action === 'song_status') {
    list = rows_('song');
    for (i = 0; i < list.length; i++) if (list[i].id === p.id) list[i]['상태'] = String(p.status || '후보');
    put_('song', list);
    return;
  }
  if (action === 'song_remove') {
    put_('song', rows_('song').filter(function (x) { return x.id !== p.id; }));
    put_('vote', rows_('vote').filter(function (x) { return x['곡'] !== p.id; }));
    var hl = rows_('hist');
    for (i = 0; i < hl.length; i++) {
      hl[i]['곡'] = splitList_(hl[i]['곡']).filter(function (x) { return x !== p.id; }).join(',');
    }
    put_('hist', hl);
    return;
  }

  /* ── 이력 ── */
  if (action === 'hist_save') {
    var h = JSON.parse(p.payload || '{}');
    if (!normDate_(h.date)) return { ok: false, error: '날짜가 필요합니다' };
    list = rows_('hist');
    var hrow = {
      id: h.id || uid_('h'), '밴드': band, '날짜': h.date,
      '시작': h.from || '', '종료': h.to || '', '합주실': h.place || '', '룸': h.room || '',
      '상태': h.status || '예정',
      '불참': (h.absent || []).join(','), '곡': (h.songs || []).join(','),
      '메모': h.memo || ''
    };
    var hh2 = false;
    for (i = 0; i < list.length; i++) if (list[i].id === hrow.id) { list[i] = hrow; hh2 = true; }
    if (!hh2) list.push(hrow);
    put_('hist', list);
    return;
  }
  if (action === 'hist_remove') {
    put_('hist', rows_('hist').filter(function (x) { return x.id !== p.id; }));
    return;
  }

  /* ── 투표 ── */
  if (action === 'vote_set') {
    var vn = String(p.name || '').trim(), vs = String(p.id || '');
    if (!vn || !vs) return { ok: false, error: '이름과 곡이 필요합니다' };
    var v = parseInt(p.value, 10) || 0;
    list = rows_('vote').filter(function (x) {
      return !(x['밴드'] === band && x['곡'] === vs && x['이름'] === vn);
    });
    if (v !== 0) list.push({ '밴드': band, '곡': vs, '이름': vn, '값': String(v) });
    put_('vote', list);
    return;
  }

  /* ── 일괄 입력 (초기 데이터 이관용) ── */
  if (action === 'bulk_import') {
    var body2 = JSON.parse(p.payload || '{}');
    if (body2.members) {
      var ml = rows_('member').filter(function (x) { return x['밴드'] !== band; });
      for (i = 0; i < body2.members.length; i++) ml.push({ '밴드': band, '이름': body2.members[i] });
      put_('member', ml);
    }
    if (body2.songs) {
      var sl = rows_('song').filter(function (x) { return x['밴드'] !== band; });
      for (i = 0; i < body2.songs.length; i++) {
        var bs = body2.songs[i];
        sl.push({ id: bs.id || uid_('s') + i, '밴드': band, '상태': bs.status || '후보',
          '제목': bs.title || '', '아티스트': bs.artist || '', '키': bs.key || '',
          '선곡자': bs.picker || '', '링크': bs.links || '', '파트': bs.parts || '',
          '튜닝': bs.tuning || '', '메모': bs.memo || '', '추가일': bs.added || today_() });
      }
      put_('song', sl);
    }
    if (body2.hist) {
      var hl2 = rows_('hist').filter(function (x) { return x['밴드'] !== band; });
      for (i = 0; i < body2.hist.length; i++) {
        var bh = body2.hist[i];
        hl2.push({ id: bh.id || uid_('h') + i, '밴드': band, '날짜': bh.date || '',
          '시작': bh.from || '', '종료': bh.to || '', '합주실': bh.place || '', '룸': bh.room || '',
          '상태': bh.status || '예정', '불참': (bh.absent || []).join(','),
          '곡': (bh.songs || []).join(','), '메모': bh.memo || '' });
      }
      put_('hist', hl2);
    }
    return;
  }

  return { ok: false, error: 'unknown action: ' + action };
}

/* ═══════════════ 엔트리 ═══════════════ */

function doGet(e) {
  var p = (e && e.parameter) || {};
  var out;
  try { out = handle_(p); }
  catch (err) { out = { ok: false, error: String(err) }; }

  var body = JSON.stringify(out);
  if (p.callback) {
    return ContentService.createTextOutput(p.callback + '(' + body + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

/** 서버 대 서버용 (curl / PowerShell). 브라우저는 doGet + JSONP 를 씁니다. */
function doPost(e) {
  var out;
  try { out = handle_(JSON.parse(e.postData.contents)); }
  catch (err) { out = { ok: false, error: String(err) }; }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}
