/**
 * THE PRIVATE EDIT 초청 회신 API (이스트엔드 영업마케팅실)
 *  - 초대장(GitHub Pages)  → doPost action=rsvp     : 참석/미참석 회신 저장
 *  - 초대장                → doGet  action=guest    : 개인 링크 코드로 이름·이전 회신 조회
 *  - 대시보드              → doGet  action=list     : 명단 + 회신 전체 (관리자 토큰)
 *  - 대시보드              → doPost action=set      : 담당자가 DM으로 받은 회신을 직접 표시 (관리자 토큰)
 *  - 이 PC 스크립트        → doPost action=import   : 명단 넣기/바꾸기 (관리자 토큰)
 * 데이터는 setup()이 만든 비공개 스프레드시트에만 저장한다(시트 ID = 스크립트 속성 SPREADSHEET_ID).
 * 관리자 토큰 원문은 코드에 없다. SHA-256 해시만 둔다.
 */
var GUEST_SHEET = '명단';
var RSVP_SHEET = '회신';
var GUEST_HEAD = ['No', '인플루언서', '팔로워', '채널', '담당자', '코드', '비고'];
var RSVP_HEAD = ['회신시각', '코드', '이름', '인스타ID', '참석', '동반', '경로'];
var ADMIN_HASH = 'c078b98e4985c204ffca201ab126a2577476f3bf1b62744ae82b93421c8908f2';
var SHEET_TITLE = 'THE PRIVATE EDIT 11.25 초청 회신 (비공개)';

/** 처음 한 번 실행: 비공개 시트 만들기(이미 있으면 탭만 점검) */
function setup() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SPREADSHEET_ID');
  var ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.create(SHEET_TITLE);
  if (!id) props.setProperty('SPREADSHEET_ID', ss.getId());
  ensureSheets_(ss);
  Logger.log('회신 시트: ' + ss.getUrl());
}

function ensureSheets_(ss) {
  var g = ss.getSheetByName(GUEST_SHEET) || ss.insertSheet(GUEST_SHEET);
  if (g.getLastRow() === 0) { g.appendRow(GUEST_HEAD); g.setFrozenRows(1); }
  var r = ss.getSheetByName(RSVP_SHEET) || ss.insertSheet(RSVP_SHEET);
  if (r.getLastRow() === 0) { r.appendRow(RSVP_HEAD); r.setFrozenRows(1); }
  ['Sheet1', '시트1'].forEach(function (n) {
    var s = ss.getSheetByName(n);
    if (s && ss.getSheets().length > 2) ss.deleteSheet(s);
  });
}

function book_() {
  var id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('setup()을 먼저 실행하세요');
  return SpreadsheetApp.openById(id);
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function isAdmin_(token) {
  if (!token) return false;
  var d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(token), Utilities.Charset.UTF_8);
  var hex = d.map(function (b) { var v = (b < 0 ? b + 256 : b).toString(16); return v.length === 1 ? '0' + v : v; }).join('');
  return hex === ADMIN_HASH;
}

/** 시트에 넣을 글자: 수식으로 해석되지 않게, 길이 제한 */
function clean_(s, max) {
  s = String(s == null ? '' : s).replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, max || 60);
  if (/^[=+\-@]/.test(s)) s = "'" + s;
  return s;
}

function handleOf_(url) {
  var m = String(url || '').match(/instagram\.com\/([A-Za-z0-9._]+)/i);
  return m ? m[1].toLowerCase() : '';
}

function guests_(ss) {
  var sh = ss.getSheetByName(GUEST_SHEET), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, GUEST_HEAD.length).getValues()
    .filter(function (r) { return String(r[1]).trim() !== ''; })
    .map(function (r) {
      return { no: r[0], name: String(r[1]), followers: String(r[2]), url: String(r[3]),
               manager: String(r[4]), code: String(r[5]).trim(), note: String(r[6]) };
    });
}

function rsvps_(ss) {
  var sh = ss.getSheetByName(RSVP_SHEET), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, RSVP_HEAD.length).getValues().map(function (r) {
    var t = r[0] instanceof Date ? r[0].toISOString() : String(r[0]);
    return { ts: t, code: String(r[1]).trim(), name: String(r[2]), insta: String(r[3]),
             attend: String(r[4]), companions: Number(r[5]) || 0, via: String(r[6]) };
  });
}

function latestFor_(list, code) {
  var last = null;
  list.forEach(function (x) { if (x.code === code) last = x; });
  return last;
}

function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    var ss = book_();
    if (p.action === 'guest') {
      var code = String(p.g || '').trim();
      if (!code) return out_({ ok: false, msg: 'no code' });
      var g = guests_(ss).filter(function (x) { return x.code === code; })[0];
      if (!g) return out_({ ok: false, msg: 'not found' });
      var last = latestFor_(rsvps_(ss), code);
      var has = last && last.attend !== 'CLEAR';
      return out_({ ok: true, name: g.name, no: g.no, rsvp: has ? { attend: last.attend, companions: last.companions } : null });
    }
    if (p.action === 'list') {
      if (!isAdmin_(p.token)) return out_({ ok: false, msg: 'auth' });
      return out_({ ok: true, guests: guests_(ss), responses: rsvps_(ss), now: new Date().toISOString() });
    }
    return out_({ ok: false, msg: 'unknown action' });
  } catch (err) {
    return out_({ ok: false, msg: String(err && err.message || err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var b = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var ss = book_();
    if (b.action === 'rsvp') return out_(saveRsvp_(ss, b, false));
    if (b.action === 'set') {
      if (!isAdmin_(b.token)) return out_({ ok: false, msg: 'auth' });
      return out_(saveRsvp_(ss, b, true));
    }
    if (b.action === 'import') {
      if (!isAdmin_(b.token)) return out_({ ok: false, msg: 'auth' });
      lock.waitLock(20000);
      return out_(importGuests_(ss, b.rows || []));
    }
    if (b.action === 'purgeTest') {
      if (!isAdmin_(b.token)) return out_({ ok: false, msg: 'auth' });
      lock.waitLock(20000);
      return out_(purgeTest_(ss));
    }
    return out_({ ok: false, msg: 'unknown action' });
  } catch (err) {
    return out_({ ok: false, msg: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function saveRsvp_(ss, b, byStaff) {
  var attend = b.attend === 'Y' ? 'Y' : b.attend === 'N' ? 'N' : b.attend === 'CLEAR' && byStaff ? 'CLEAR' : '';
  if (!attend) return { ok: false, msg: '참석 여부를 골라 주세요.' };
  var comp = attend === 'Y' && Number(b.companions) === 1 ? 1 : 0;
  var guests = guests_(ss), code = String(b.g || '').trim(), name = '', insta = '', via = '';
  if (code) {
    var g = guests.filter(function (x) { return x.code === code; })[0];
    if (!g) return { ok: false, msg: '초대 링크를 다시 확인해 주세요.' };
    name = g.name; via = byStaff ? '담당자 입력' : '개인 링크';
  } else {
    name = clean_(b.name, 40);
    insta = clean_(String(b.insta || '').replace(/^@/, ''), 40).toLowerCase();
    if (!name) return { ok: false, msg: '이름을 적어 주세요.' };
    if (!/^[a-z0-9._]{2,30}$/.test(insta)) return { ok: false, msg: '인스타그램 ID를 확인해 주세요.' };
    var hit = guests.filter(function (x) { return handleOf_(x.url) === insta; })[0];
    if (hit) { code = hit.code; via = '직접 입력(명단 일치)'; } else via = '직접 입력';
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    ss.getSheetByName(RSVP_SHEET).appendRow([new Date(), code, name, insta, attend, comp, via]);
    SpreadsheetApp.flush();
  } finally { lock.releaseLock(); }
  return { ok: true, matched: !!code };
}

/** 배포 확인용 시험 회신(이름이 ZZTEST로 시작)만 지운다 */
function purgeTest_(ss) {
  var sh = ss.getSheetByName(RSVP_SHEET), n = sh.getLastRow(), del = 0;
  for (var r = n; r >= 2; r--) {
    if (/^'?ZZTEST/.test(String(sh.getRange(r, 3).getValue()))) { sh.deleteRow(r); del++; }
  }
  return { ok: true, deleted: del };
}

function importGuests_(ss, rows) {
  if (!rows.length) return { ok: false, msg: 'no rows' };
  var sh = ss.getSheetByName(GUEST_SHEET);
  var vals = rows.map(function (r) {
    return [r.no || '', clean_(r.name, 40), clean_(r.followers, 20), clean_(r.url, 200),
            clean_(r.manager, 20), clean_(r.code, 20), clean_(r.note, 100)];
  });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, GUEST_HEAD.length).clearContent();
  sh.getRange(2, 1, vals.length, GUEST_HEAD.length).setValues(vals);
  SpreadsheetApp.flush();
  return { ok: true, count: vals.length };
}
