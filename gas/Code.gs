/**
 * THE PRIVATE EDIT 초청 회신 API (이스트엔드 영업마케팅실)
 *  - 초대장(GitHub Pages)  → doPost action=rsvp     : 참석/미참석 회신 저장
 *  - 초대장                → doGet  action=guest    : 개인 링크 코드로 이름·이전 회신 조회
 *  - 대시보드              → doGet  action=list     : 명단 + 회신 전체 (관리자 토큰)
 *  - 대시보드              → doPost action=set      : 담당자가 DM으로 받은 회신을 직접 표시 (관리자 토큰)
 *  - 이 PC 스크립트        → doPost action=import   : 명단 넣기/바꾸기 (관리자 토큰) + 마케팅팀 시트에 개인 링크 쓰기
 * 회신(참석/미참석·동반)은 마케팅팀 시트 '인플루언서 초청' 표의 「참석 여부」「참여 인원」에도 바로 적는다.
 * 데이터는 setup()이 만든 비공개 스프레드시트에만 저장한다(시트 ID = 스크립트 속성 SPREADSHEET_ID).
 * 관리자 토큰 원문은 코드에 없다. SHA-256 해시만 둔다.
 */
var GUEST_SHEET = '명단';
var RSVP_SHEET = '회신';
var GUEST_HEAD = ['No', '인플루언서', '팔로워', '채널', '담당자', '코드', '비고'];
var RSVP_HEAD = ['회신시각', '코드', '이름', '인스타ID', '참석', '동반', '경로'];
var ADMIN_HASH = 'c078b98e4985c204ffca201ab126a2577476f3bf1b62744ae82b93421c8908f2';
var SHEET_TITLE = 'THE PRIVATE EDIT 11.25 초청 회신 (비공개)';
var MKT_ID = '1XoLURnQZmwYM-TuB8hJfuIjNiM9LWPI3J9-aaDPDMKM';   // 시티브리즈_마케팅팀
var MKT_GID = 2033474354;                                        // '인플루언서 초청' 탭
var MKT_TAB = '인플루언서 초청';

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
      return out_(importGuests_(ss, b.rows || [], String(b.base || '')));
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
  var synced = false;
  if (code) { try { synced = syncMkt_(guests, code, attend, comp); } catch (x) { Logger.log('마케팅 시트 반영 실패: ' + x); } }
  return { ok: true, matched: !!code, synced: synced };
}

/** 배포 확인용 시험 회신(이름이 ZZTEST로 시작)만 지운다 */
function purgeTest_(ss) {
  var sh = ss.getSheetByName(RSVP_SHEET), n = sh.getLastRow(), del = 0;
  for (var r = n; r >= 2; r--) {
    if (/^'?ZZTEST/.test(String(sh.getRange(r, 3).getValue()))) { sh.deleteRow(r); del++; }
  }
  return { ok: true, deleted: del };
}

function importGuests_(ss, rows, base) {
  if (!rows.length) return { ok: false, msg: 'no rows' };
  var sh = ss.getSheetByName(GUEST_SHEET);
  var vals = rows.map(function (r) {
    return [r.no || '', clean_(r.name, 40), clean_(r.followers, 20), clean_(r.url, 200),
            clean_(r.manager, 20), clean_(r.code, 20), clean_(r.note, 100)];
  });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, GUEST_HEAD.length).clearContent();
  sh.getRange(2, 1, vals.length, GUEST_HEAD.length).setValues(vals);
  SpreadsheetApp.flush();
  var links = 0;
  if (base) links = writeLinks_(rows, base);
  return { ok: true, count: vals.length, links: links };
}

/* ---------- 마케팅팀 시트 '인플루언서 초청' 표 연동 ---------- */

/** 머리줄(「인플루언서」와 「초대장 링크」가 같이 있는 줄)을 찾아 열 위치를 돌려준다 */
function mkt_() {
  var ss = SpreadsheetApp.openById(MKT_ID);
  var sh = ss.getSheets().filter(function (x) { return x.getSheetId() === MKT_GID; })[0] || ss.getSheetByName(MKT_TAB);
  var nr = Math.min(sh.getLastRow(), 400), nc = Math.min(sh.getLastColumn(), 40);
  var vals = sh.getRange(1, 1, nr, nc).getDisplayValues();
  for (var r = 0; r < nr; r++) {
    var row = vals[r].map(function (v) { return String(v).replace(/\s+/g, ''); });
    var link = row.indexOf('초대장링크'), name = row.indexOf('인플루언서');
    if (link >= 0 && name >= 0) {
      var no = row.indexOf('No.'); if (no < 0) no = row.indexOf('No');
      return { sh: sh, vals: vals, head: r, col: { no: no, name: name, attend: row.indexOf('참석여부'), people: row.indexOf('참여인원'), link: link } };
    }
  }
  throw new Error('마케팅 시트에서 「초대장 링크」 머리줄을 못 찾음');
}

/** 명단 번호(No.)로 표의 줄(1부터)을 찾는다. 링크 칸에 코드가 있으면 그것을 먼저 본다 */
function mktRow_(m, no, code) {
  for (var r = m.head + 1; r < m.vals.length; r++) {
    if (code && String(m.vals[r][m.col.link]).indexOf('g=' + code) >= 0) return r + 1;
  }
  for (var r2 = m.head + 1; r2 < m.vals.length; r2++) {
    if (m.col.no >= 0 && String(m.vals[r2][m.col.no]).trim() === String(no) && String(m.vals[r2][m.col.name]).trim()) return r2 + 1;
  }
  return 0;
}

/** 칸에 목록형 입력 규칙이 있으면 그 목록 안의 값으로 맞춘다 */
function fitList_(cell, want, re) {
  var v = cell.getDataValidation();
  if (!v || v.getCriteriaType() !== SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) return want;
  var list = v.getCriteriaValues()[0] || [];
  if (list.indexOf(want) >= 0) return want;
  for (var i = 0; i < list.length; i++) if (re.test(String(list[i]))) return list[i];
  return want;
}

function writeLinks_(rows, base) {
  var m = mkt_(), n = 0;
  rows.forEach(function (g) {
    if (!g.code) return;
    var r = mktRow_(m, g.no, g.code);
    if (!r) return;
    m.sh.getRange(r, m.col.link + 1).setValue(base + (base.indexOf('?') >= 0 ? '&' : '?') + 'g=' + g.code);
    n++;
  });
  SpreadsheetApp.flush();
  return n;
}

function syncMkt_(guests, code, attend, comp) {
  var g = guests.filter(function (x) { return x.code === code; })[0];
  if (!g) return false;
  var m = mkt_(), r = mktRow_(m, g.no, code);
  if (!r) return false;
  if (m.col.attend >= 0) {
    var c = m.sh.getRange(r, m.col.attend + 1);
    if (attend === 'CLEAR') c.clearContent();
    else c.setValue(attend === 'Y' ? fitList_(c, '참석', /^(참석|O|○|Y|예)/i) : fitList_(c, '미참석', /^(미참석|불참|X|×|N|아니)/i));
  }
  if (m.col.people >= 0) {
    var p = m.sh.getRange(r, m.col.people + 1);
    if (attend === 'CLEAR') p.clearContent(); else p.setValue(attend === 'Y' ? 1 + comp : 0);
  }
  SpreadsheetApp.flush();
  return true;
}

/** 편집기에서 한 번 눌러 보는 점검용: 마케팅 시트 머리줄·열 위치를 로그로 */
function checkMkt() {
  var m = mkt_();
  Logger.log('머리줄 ' + (m.head + 1) + '행, 열(0부터) ' + JSON.stringify(m.col));
}
