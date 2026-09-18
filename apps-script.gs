/**
 * 밴드매니저 — Google Apps Script 웹앱 (v3)
 *
 * 스프레드시트는 "데이터 저장소"일 뿐입니다. 모든 조작은 웹페이지에서 합니다.
 * 탭 8개 (스크립트가 알아서 만듭니다). 모든 행은 첫 컬럼에 밴드를 답니다.
 *
 *   밴드   id | 이름
 *   멤버   밴드 | 이름
 *   설정   밴드 | key | value           (dates / hourStart / hourEnd / title)
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

function ss_() { return SpreadsheetApp.openById(SHEET_ID); }

/** 헤더가 맞는 탭을 보장한다. 옛 탭은 이름만 바꿔 보존. */
function tab_(key) {
  var spec = TABS[key], ss = ss_(), sh = ss.getSheetByName(spec.name);

  if (sh) {
    var lastCol = Math.max(sh.getLastColumn(), 1);
    var head = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    var same = true;
    for (var i = 0; i < spec.head.length; i++) {
      if (String(head[i] || '').trim() !== spec.head[i]) { same = false; break; }
    }
    if (same) return sh;
    var bak = spec.name + '_구버전';
    if (ss.getSheetByName(bak)) bak += '_' + Utilities.formatDate(new Date(), 'Asia/Seoul', 'MMddHHmm');
    sh.setName(bak);
    sh = null;
  }

  sh = ss.insertSheet(spec.name);
  sh.getRange(1, 1, 1, spec.head.length).setValues([spec.head]).setFontWeight('bold');
  sh.setFrozenRows(1);
  return sh;
}

/** 탭 전체를 객체 배열로 */
function rows_(key) {
  var spec = TABS[key], sh = tab_(key);
  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  var vals = sh.getRange(2, 1, lastRow - 1, spec.head.length).getValues();
  var out = [];
  for (var r = 0; r < vals.length; r++) {
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
  return out;
}

/** 탭 전체 덮어쓰기 */
function put_(key, list) {
  var spec = TABS[key], sh = tab_(key);
  var last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, spec.head.length).clearContent();
  if (!list.length) return;
  var out = [];
  for (var r = 0; r < list.length; r++) {
    var row = [];
    for (var c = 0; c < spec.head.length; c++) {
      var v = list[r][spec.head[c]];
      row.push(v === null || v === undefined ? '' : String(v));
    }
    out.push(row);
  }
  sh.getRange(2, 1, out.length, spec.head.length).setNumberFormat('@').setValues(out);
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

/** "10,11,12" → [10,11,12] / "-" → [] / 빈칸 → null(미응답) */
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

function state_() {
  var bl = bands_();
  var out = {
    ok: true,
    bands: [], members: {}, config: {}, hours: {}, notes: {},
    songs: [], hist: [], votes: {},
    now: Utilities.formatDate(new Date(), 'Asia/Seoul', "yyyy-MM-dd'T'HH:mm:ss")
  };

  for (var i = 0; i < bl.length; i++) {
    var b = bl[i].id;
    out.bands.push({ id: b, name: bl[i]['이름'] });
    out.members[b] = [];
    out.config[b] = { dates: [], hourStart: CONF_DEFAULT.hourStart,
                      hourEnd: CONF_DEFAULT.hourEnd, title: bl[i]['이름'] };
    out.hours[b] = {};
    out.notes[b] = {};
  }

  var ms = rows_('member'), i2;
  for (i2 = 0; i2 < ms.length; i2++) {
    var mb = ms[i2]['밴드'];
    if (out.members[mb]) out.members[mb].push(ms[i2]['이름']);
  }

  var cs = rows_('conf');
  for (i2 = 0; i2 < cs.length; i2++) {
    var cb = cs[i2]['밴드'], k = cs[i2]['key'], v = cs[i2]['value'];
    if (!out.config[cb]) continue;
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
  }

  var rs = rows_('resp');
  for (i2 = 0; i2 < rs.length; i2++) {
    var rb = rs[i2]['밴드'], rn = rs[i2]['이름'], rd = normDate_(rs[i2]['날짜']);
    if (!out.hours[rb] || !rn || !rd) continue;
    var hh = parseHours_(rs[i2]['시간']);
    if (hh === null) continue;
    if (!out.hours[rb][rn]) out.hours[rb][rn] = {};
    out.hours[rb][rn][rd] = hh;
  }

  var ns = rows_('note');
  for (i2 = 0; i2 < ns.length; i2++) {
    var nb = ns[i2]['밴드'], nn = ns[i2]['이름'], ndd = normDate_(ns[i2]['날짜']);
    if (!out.notes[nb] || !nn || !ndd || !ns[i2]['메모']) continue;
    if (!out.notes[nb][nn]) out.notes[nb][nn] = {};
    out.notes[nb][nn][ndd] = ns[i2]['메모'];
  }

  var ss = rows_('song');
  for (i2 = 0; i2 < ss.length; i2++) {
    var s = ss[i2];
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
    if (!vsid || !vn) continue;
    if (!out.votes[vsid]) out.votes[vsid] = {};
    out.votes[vsid][vn] = parseInt(vs[i2]['값'], 10) || 0;
  }

  return out;
}

/* ═══════════════ 동작 ═══════════════ */

function handle_(p) {
  var action = p.action || 'load';
  if (action === 'load') return state_();

  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    var r = act_(action, p);
    if (r && r.ok === false) return r;
    SpreadsheetApp.flush();
    return state_();
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
