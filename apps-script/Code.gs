/**
 * ============================================================
 *  경남2사무소 영업활동 대시보드 v2 — 데이터 API (Google Apps Script)
 * ============================================================
 *  - 비밀번호: [프로젝트 설정] → [스크립트 속성] 에 OFFICE_PASSWORD 추가
 *  - 데이터 시트
 *      방문기록 : 방문 1건 = 1행 (웹앱 입력 폼이 여기에 쌓음)
 *      거래처   : 거래처 마스터 (담당자, 집중품목, 성향태그 등)
 *      품목     : 제안 품목 목록 (+ 자동 인식용 별칭)
 *      담당자   : 직원 목록 (+ 평가실적 탭의 이전 담당자 이름)
 *      N월 평가실적 / 27년 N월 평가실적 : 매월 본사 실적을 붙여넣는 탭 (기존 방식 유지)
 */

var SHEETS = { VISIT: '방문기록', ACCOUNT: '거래처', PRODUCT: '품목', STAFF: '담당자' };

var VISIT_HEADERS = ['ID', '방문일', '방문자', '거래처코드', '거래처명', '제안품목', '제안내용', '반응',
  '결과', '다음액션', '다음예정일', '동행', '입력일시', '출처', '비고'];
/** 구분: 방문처 / DCM(도매 매출자료로 평가 — 방문 통계에서 제외)
 *  거래처코드가 아직 없는 신규처는 TMP-0001 같은 임시코드를 쓰고, 실적이 나오면 "실적 연결"로 진짜 코드로 바꿉니다. */
var ACCOUNT_HEADERS = ['거래처코드', '거래처명', '담당자', '구분', '진료과', '집중품목', '성향태그', '메모', '사용여부', '추가코드'];
/* 추가코드: 같은 거래처가 평가실적에 다른 코드로도 잡힐 때 (쉼표로 구분) — 실적을 주 코드에 합산 */
var ACCOUNT_TYPES = ['방문처', 'DCM'];
var PRODUCT_HEADERS = ['품목명', '별칭(쉼표로 구분)', '사용여부'];
var STAFF_HEADERS = ['이름', '사용여부', '이전 담당자 이름(쉼표로 구분)'];

var REACTIONS = ['긍정', '보류', '거절'];
var COMPANIONS = ['CM', 'PM', '소장'];

/** "N월 평가실적" 처럼 연도가 없는 탭의 연도 */
var DEFAULT_SALES_YEAR = 2026;
var TOKEN_SECONDS = 6 * 3600;

var DEFAULT_STAFF = [
  ['이수빈', 'Y', '구원모'],
  ['김준영', 'Y', ''],
  ['유승태', 'Y', ''],
  ['장문수', 'Y', ''],
  ['정재은', 'Y', ''],
  ['고현민', 'Y', '박준수, 계상현']
];

/** [품목명, 별칭] — 품목 시트에서 언제든 수정 가능 */
var DEFAULT_PRODUCTS = [
  ['펙수클루', '펙, 팩수'], ['바로에젯', '바로에, 바/, /바'], ['실리스칸', '실리, 실/, /실'], ['이스미젠', '이스미'],
  ['웰체크', '웰만사, 웰체'], ['일만사', '웰만사'], ['이지에프', '이지에프MD'], ['옵티나', '옵티'],
  ['페라미빅트', ''], ['카트BP', ''], ['클린콜', ''], ['웨이메드 엔도', '웨이메드'],
  ['우루사', ''], ['스토보클로', '스토보'], ['엔블로', ''], ['나보타', ''], ['브리바탑', ''],
  ['코그테라', ''], ['루피어데포', '루피어'], ['오스테오시그널', ''], ['압솔라', ''],
  ['올메텍', '올메'], ['크레젯', ''], ['푸루나졸', ''], ['제미메트', ''], ['스피로킷', ''],
  ['닥터빌', ''], ['이지듀', ''], ['리브레', ''], ['레가론', ''], ['피타바', ''],
  ['글리아타민', ''], ['에티아', ''], ['thynC', '씽크'], ['e-DToD', '']
];

/* =========================== API 진입 =========================== */
/*  화면(index.html)은 GitHub Pages 에 있고, 이 스크립트는 데이터 창구(API) 역할만 합니다.
 *  화면 → POST { fn, token, args } → { ok, data | error }                                    */

var API_ = {
  login: login,
  getBootstrap: getBootstrap,
  getDashboard: getDashboard,
  getAccountDetail: getAccountDetail,
  saveVisit: saveVisit,
  updateAccountProfile: updateAccountProfile,
  addAccount: addAccount,
  linkAccountCode: linkAccountCode,
  mergeAccountCode: mergeAccountCode
};

function doPost(e) {
  var out;
  try {
    var req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    var fn = Object.prototype.hasOwnProperty.call(API_, req.fn) ? API_[req.fn] : null;
    if (!fn) throw new Error('알 수 없는 요청입니다.');
    var args = Array.isArray(req.args) ? req.args : [];
    var data = req.fn === 'login' ? fn(args[0]) : fn.apply(null, [req.token].concat(args));
    out = { ok: true, data: data === undefined ? null : data };
  } catch (err) {
    out = { ok: false, error: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/** 브라우저로 주소를 직접 열었을 때: 작동 확인용 */
function doGet() {
  return ContentService.createTextOutput(JSON.stringify({ ok: true, service: '경남2사무소 영업활동 API' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/* =========================== 인증 =========================== */

function login(password) {
  var real = PropertiesService.getScriptProperties().getProperty('OFFICE_PASSWORD');
  if (!real) throw new Error('비밀번호가 아직 설정되지 않았습니다. 관리자에게 문의하세요.');
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('login_fails') || 0);
  if (fails >= 10) throw new Error('로그인 실패가 많아 10분 동안 잠겼습니다.');
  if (String(password) !== String(real)) {
    cache.put('login_fails', String(fails + 1), 600);
    throw new Error('비밀번호가 올바르지 않습니다.');
  }
  var token = Utilities.getUuid();
  cache.put('tok_' + token, '1', TOKEN_SECONDS);
  return token;
}

function requireAuth_(token) {
  var cache = CacheService.getScriptCache();
  if (!token || !cache.get('tok_' + token)) throw new Error('AUTH_EXPIRED');
  cache.put('tok_' + token, '1', TOKEN_SECONDS); // 사용 중이면 연장
}

/* =========================== 웹앱 API =========================== */

/** 처음 접속 시: 입력 폼과 필터에 필요한 기본 목록 */
function getBootstrap(token) {
  requireAuth_(token);
  var staff = getStaff_();
  return {
    staff: staff.names,
    products: getProducts_().filter(function (p) { return p.active; }).map(function (p) { return p.name; }),
    accounts: getAccounts_().map(function (a) { return { code: a.code, name: a.name, owner: a.owner, type: a.type }; }),
    reactions: REACTIONS,
    companions: COMPANIONS,
    today: fmtDate_(new Date())
  };
}

/** 대시보드 전체 데이터 (방문 본문 제외 — 상세는 getAccountDetail) */
function getDashboard(token) {
  requireAuth_(token);
  var staff = getStaff_();
  var accounts = getAccounts_();
  var sales = readSales_(staff.alias);
  var visits = readVisits_();

  var byCode = {};
  accounts.forEach(function (a) {
    byCode[a.code] = {
      code: a.code, name: a.name, owner: a.owner, type: a.type, dept: a.dept, focus: a.focus,
      tags: a.tags, extraCodes: a.extraCodes, sales: accountSales_(a, sales.byCode), registered: true
    };
  });
  var used = usedCodes_(accounts);
  // 평가실적에는 있지만 거래처 시트에 없는 곳 (합계가 본사 자료와 맞도록 포함)
  Object.keys(sales.byCode).forEach(function (code) {
    if (used[code]) return;
    var info = sales.info[code] || {};
    byCode[code] = {
      code: code, name: info.name || code, owner: info.owner || '(미지정)', type: '방문처', dept: '', focus: '',
      tags: '', sales: sales.byCode[code], registered: false
    };
  });

  var light = [];
  var last = {};
  visits.forEach(function (v) {
    light.push([v.date, v.visitor, v.code, v.reaction]);
    var cur = last[v.code];
    if (!cur || v.date >= cur.date) {
      last[v.code] = {
        date: v.date, visitor: v.visitor, reaction: v.reaction,
        result: String(v.result).slice(0, 140), nextAction: v.nextAction, nextDate: v.nextDate
      };
    }
  });

  return {
    accounts: Object.keys(byCode).map(function (k) { return byCode[k]; }),
    visits: light,                 // [date, visitor, code, reaction]
    lastVisit: last,
    staff: staff.names,
    warnings: sales.warnings,
    generatedAt: Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm')
  };
}

/** 거래처 1곳의 전체 히스토리 */
function getAccountDetail(token, code) {
  requireAuth_(token);
  code = String(code);
  var acc = getAccounts_().filter(function (a) { return a.code === code; })[0] || null;
  var visits = readVisits_().filter(function (v) { return v.code === code; });
  visits.sort(function (a, b) { return a.date < b.date ? 1 : a.date > b.date ? -1 : 0; });
  var sales = readSales_(getStaff_().alias);
  return { account: acc, visits: visits, sales: acc ? accountSales_(acc, sales.byCode) : (sales.byCode[code] || {}) };
}

/** 방문 기록 저장 */
function saveVisit(token, p) {
  requireAuth_(token);
  var staff = getStaff_().names;
  var accounts = getAccounts_();
  var acc = accounts.filter(function (a) { return a.code === String(p.code); })[0];

  var date = parseDate_(p.date);
  if (!date) throw new Error('방문일을 확인해 주세요.');
  if (staff.indexOf(p.visitor) === -1) throw new Error('방문자를 선택해 주세요.');
  if (!acc) throw new Error('거래처를 선택해 주세요.');
  if (REACTIONS.indexOf(p.reaction) === -1) throw new Error('원장님 반응을 선택해 주세요.');
  if (!String(p.result || '').trim()) throw new Error('결과(원장님 코멘트)를 입력해 주세요.');
  var nextDate = p.nextDate ? parseDate_(p.nextDate) : '';

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = sheet_(SHEETS.VISIT);
    var id = 'V' + Utilities.formatDate(new Date(), tz_(), 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10);
    sh.appendRow([
      id, date, p.visitor, acc.code, acc.name,
      (p.products || []).join(', '), clean_(p.proposal), p.reaction, clean_(p.result),
      clean_(p.nextAction), nextDate || '', (p.companions || []).join(', '),
      new Date(), '웹입력', ''
    ]);
    var r = sh.getLastRow();
    sh.getRange(r, 2).setNumberFormat('yyyy-mm-dd');
    sh.getRange(r, 11).setNumberFormat('yyyy-mm-dd');
    return id;
  } finally {
    lock.releaseLock();
  }
}

/** 거래처 프로필(구분·진료과·성향태그·메모) 수정 */
function updateAccountProfile(token, code, profile) {
  requireAuth_(token);
  var t = readTable_(SHEETS.ACCOUNT);
  var row = t.rows.filter(function (r) { return normalizeCode_(r['거래처코드']) === String(code); })[0];
  if (!row) throw new Error('거래처를 찾을 수 없습니다.');
  var set = function (h, v) {
    var c = t.headers.indexOf(h);
    if (c !== -1) t.sheet.getRange(row._row, c + 1).setValue(clean_(v));
  };
  if (ACCOUNT_TYPES.indexOf(profile.type) !== -1) set('구분', profile.type);
  set('진료과', profile.dept);
  set('성향태그', profile.tags);
  set('메모', profile.memo);
  return true;
}

/**
 * 거래처 추가 (방문 입력 화면의 "+ 신규 거래처", 또는 평가실적에만 있는 곳 등록)
 * 거래처코드를 모르면 비워두면 TMP-0001 형식의 임시코드가 붙습니다.
 */
function addAccount(token, p) {
  requireAuth_(token);
  var name = clean_(p.name);
  if (!name) throw new Error('거래처명을 입력해 주세요.');
  if (getStaff_().names.indexOf(p.owner) === -1) throw new Error('담당자를 선택해 주세요.');
  var code = '';
  if (String(p.code || '').trim()) {
    code = normalizeCode_(p.code);
    if (!code || isTmpCode_(code)) throw new Error('거래처코드는 숫자만 입력해 주세요. 모르면 비워두세요.');
  }
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var t = readTable_(SHEETS.ACCOUNT);
    var codes = t.rows.map(function (r) { return normalizeCode_(r['거래처코드']); });
    if (code && usedCodes_(getAccounts_(true))[code]) throw new Error('이미 등록된 거래처코드입니다. (추가코드 포함)');
    var dupName = t.rows.some(function (r) { return String(r['거래처명']).trim() === name && String(r['사용여부']).trim().toUpperCase() !== 'N'; });
    if (dupName) throw new Error('같은 이름의 거래처가 이미 있습니다. 신규처에 실적을 붙이려면 거래처 상세 화면의 "실적 연결"을 이용하세요.');
    if (!code) code = nextTmpCode_(codes);
    var type = p.type === 'DCM' ? 'DCM' : '방문처';
    var values = { '거래처코드': code, '거래처명': name, '담당자': p.owner, '구분': type, '사용여부': 'Y' };
    t.sheet.appendRow(t.headers.map(function (h) { return values[h] || ''; }));
    t.sheet.getRange(t.sheet.getLastRow(), t.headers.indexOf('거래처코드') + 1).setNumberFormat('@').setValue(code);
    return { code: code, name: name, owner: p.owner, type: type };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 임시코드(TMP-0001) 신규처에 실적이 나오기 시작하면 진짜 거래처코드로 바꿔 실적을 이어붙임.
 * 거래처 시트의 코드와, 방문기록에 남아있는 같은 임시코드를 모두 바꿉니다.
 */
function linkAccountCode(token, tmpCode, realCode) {
  requireAuth_(token);
  tmpCode = String(tmpCode);
  realCode = normalizeCode_(realCode);
  if (!isTmpCode_(tmpCode)) throw new Error('임시코드 거래처만 연결할 수 있습니다.');
  if (!realCode || isTmpCode_(realCode)) throw new Error('연결할 거래처코드(숫자)를 입력해 주세요.');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var t = readTable_(SHEETS.ACCOUNT);
    var row = t.rows.filter(function (r) { return normalizeCode_(r['거래처코드']) === tmpCode; })[0];
    if (!row) throw new Error('거래처를 찾을 수 없습니다.');
    var taken = !!usedCodes_(getAccounts_(true))[realCode];
    if (taken) throw new Error('이 거래처코드는 이미 거래처 시트에 있습니다.');
    t.sheet.getRange(row._row, t.headers.indexOf('거래처코드') + 1).setNumberFormat('@').setValue(realCode);

    var vs = sheet_(SHEETS.VISIT);
    var moved = 0;
    if (vs.getLastRow() > 1) {
      var col = vs.getRange(2, VISIT_HEADERS.indexOf('거래처코드') + 1, vs.getLastRow() - 1, 1);
      var vals = col.getValues().map(function (r) {
        if (normalizeCode_(r[0]) === tmpCode) { moved++; return [realCode]; }
        return r;
      });
      if (moved) col.setValues(vals);
    }
    return { code: realCode, movedVisits: moved };
  } finally {
    lock.releaseLock();
  }
}

/* =========================== 데이터 읽기 =========================== */

function getStaff_() {
  var t = readTable_(SHEETS.STAFF);
  var names = [], alias = {};
  t.rows.forEach(function (r) {
    var n = String(r[STAFF_HEADERS[0]] || '').trim();
    if (!n) return;
    alias[n] = n;
    String(r[STAFF_HEADERS[2]] || '').split(',').forEach(function (old) {
      old = old.trim();
      if (old) alias[old] = n;
    });
    if (String(r[STAFF_HEADERS[1]]).trim().toUpperCase() !== 'N') names.push(n);
  });
  return { names: names, alias: alias };
}

function getAccounts_(includeInactive) {
  return readTable_(SHEETS.ACCOUNT).rows.map(function (r) {
    return {
      code: normalizeCode_(r['거래처코드']), name: String(r['거래처명'] || '').trim(),
      owner: String(r['담당자'] || '').trim(), dept: String(r['진료과'] || ''),
      type: String(r['구분'] || '').trim() === 'DCM' ? 'DCM' : '방문처',
      focus: String(r['집중품목'] || ''),
      tags: String(r['성향태그'] || ''), memo: String(r['메모'] || ''),
      active: String(r['사용여부']).trim().toUpperCase() !== 'N',
      extraCodes: String(r['추가코드'] || '').split(',').map(normalizeCode_).filter(String)
    };
  }).filter(function (a) { return a.code && (a.active || includeInactive); });
}

/** 주 코드 + 추가코드의 월별 실적 합산 */
function accountSales_(acc, byCode) {
  var out = {};
  [acc.code].concat(acc.extraCodes || []).forEach(function (c) {
    var s = byCode[c] || {};
    Object.keys(s).forEach(function (ym) { out[ym] = (out[ym] || 0) + s[ym]; });
  });
  return out;
}

/** 거래처 시트에서 이미 쓰고 있는 코드 (주 코드 + 추가코드) */
function usedCodes_(accounts) {
  var m = {};
  accounts.forEach(function (a) { m[a.code] = a.code; (a.extraCodes || []).forEach(function (c) { m[c] = a.code; }); });
  return m;
}

/** 평가실적에만 있는 코드를 기존 거래처의 추가코드로 합치기 (같은 병원이 코드 2개인 경우) */
function mergeAccountCode(token, mainCode, extraCode) {
  requireAuth_(token);
  extraCode = normalizeCode_(extraCode);
  if (!extraCode || isTmpCode_(extraCode)) throw new Error('합칠 거래처코드를 확인해 주세요.');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (usedCodes_(getAccounts_(true))[extraCode]) throw new Error('이 코드는 이미 거래처 시트에 등록되어 있습니다.');
    var t = readTable_(SHEETS.ACCOUNT);
    var row = t.rows.filter(function (r) { return normalizeCode_(r['거래처코드']) === String(mainCode); })[0];
    if (!row) throw new Error('합칠 대상 거래처를 찾을 수 없습니다.');
    var c = t.headers.indexOf('추가코드');
    if (c === -1) throw new Error("거래처 시트에 '추가코드' 열이 없습니다. 관리자가 setupSheets() 를 한 번 더 실행해 주세요.");
    var cur = String(row['추가코드'] || '').split(',').map(function (x) { return x.trim(); }).filter(String);
    cur.push(extraCode);
    t.sheet.getRange(row._row, c + 1).setNumberFormat('@').setValue(cur.join(', '));
    return true;
  } finally {
    lock.releaseLock();
  }
}

function getProducts_() {
  return readTable_(SHEETS.PRODUCT).rows.map(function (r) {
    return { name: String(r[PRODUCT_HEADERS[0]] || '').trim(), active: String(r[PRODUCT_HEADERS[2]]).trim().toUpperCase() !== 'N' };
  }).filter(function (p) { return p.name; });
}

function readVisits_() {
  return readTable_(SHEETS.VISIT).rows.map(function (r) {
    return {
      id: String(r['ID']), date: fmtDate_(r['방문일']), visitor: String(r['방문자'] || ''),
      code: normalizeCode_(r['거래처코드']), name: String(r['거래처명'] || ''),
      products: String(r['제안품목'] || ''), proposal: String(r['제안내용'] || ''),
      reaction: String(r['반응'] || ''), result: String(r['결과'] || ''),
      nextAction: String(r['다음액션'] || ''), nextDate: fmtDate_(r['다음예정일']),
      companions: String(r['동행'] || ''), source: String(r['출처'] || ''), note: String(r['비고'] || '')
    };
  }).filter(function (v) { return v.date && v.code; });
}

/**
 * "N월 평가실적" / "27년 N월 평가실적" 탭을 모두 읽어 거래처코드별 월 실적으로 합산
 *  - 같은 거래처코드의 서로 다른 시퀀스는 합산
 *  - 같은 거래처코드+시퀀스가 두 번 나오면 첫 번째 값만 쓰고 경고 표시 (기존 시트와 동일한 방식)
 */
function readSales_(alias) {
  var byCode = {}, info = {}, warnings = [];
  var re = /^\s*(?:(\d{2,4})\s*년\s*)?(\d{1,2})\s*월\s*평가실적\s*$/;
  SpreadsheetApp.getActiveSpreadsheet().getSheets().forEach(function (sh) {
    var m = sh.getName().match(re);
    if (!m) return;
    var year = m[1] ? Number(m[1].length === 2 ? '20' + m[1] : m[1]) : DEFAULT_SALES_YEAR;
    var ym = year + '-' + pad2_(m[2]);
    var vals = sh.getDataRange().getValues();
    if (vals.length < 2) return;
    var h = vals[0].map(function (x) { return String(x).replace(/\s+/g, ''); });
    var find = function (pred) { for (var i = 0; i < h.length; i++) if (pred(h[i])) return i; return -1; };
    var cMgr = find(function (x) { return x === '담당자'; });
    var cCode = find(function (x) { return x === '거래처코드'; });
    var cSeq = find(function (x) { return x === '시퀀스'; });
    var cName = find(function (x) { return x === '거래처명'; });
    var cVal = find(function (x) { return x.indexOf('최종실적') === 0; });
    if (cCode < 0 || cVal < 0) { warnings.push(sh.getName() + ': "거래처코드" 또는 "최종실적" 열을 찾지 못했습니다.'); return; }
    var seen = {};
    for (var r = 1; r < vals.length; r++) {
      var code = normalizeCode_(vals[r][cCode]);
      if (!code) continue;
      var v = toNumber_(vals[r][cVal]);
      var key = code + '&' + (cSeq >= 0 ? vals[r][cSeq] : '');
      if (seen[key]) {
        warnings.push(sh.getName() + ': ' + (vals[r][cName] || code) + ' (' + code + ') 이(가) 중복으로 들어 있어 첫 번째 값만 반영했습니다.');
        continue;
      }
      seen[key] = true;
      byCode[code] = byCode[code] || {};
      byCode[code][ym] = (byCode[code][ym] || 0) + v;
      var mgr = cMgr >= 0 ? String(vals[r][cMgr] || '').trim() : '';
      if (!info[code] || ym >= info[code].ym) {
        info[code] = { ym: ym, name: cName >= 0 ? String(vals[r][cName] || '').trim() : '', owner: alias[mgr] || mgr };
      }
    }
  });
  return { byCode: byCode, info: info, warnings: warnings };
}

/* =========================== 초기 설정 (편집기에서 1회 실행) =========================== */

function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ensure = function (name, headers) {
    var sh = ss.getSheetByName(name);
    var created = false;
    if (!sh) { sh = ss.insertSheet(name); created = true; }
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    }
    return { sheet: sh, created: created };
  };

  ensure(SHEETS.VISIT, VISIT_HEADERS);

  var staff = ensure(SHEETS.STAFF, STAFF_HEADERS);
  if (staff.sheet.getLastRow() === 1) staff.sheet.getRange(2, 1, DEFAULT_STAFF.length, 3).setValues(DEFAULT_STAFF);

  var prod = ensure(SHEETS.PRODUCT, PRODUCT_HEADERS);
  if (prod.sheet.getLastRow() === 1) {
    prod.sheet.getRange(2, 1, DEFAULT_PRODUCTS.length, 3).setValues(DEFAULT_PRODUCTS.map(function (p) { return [p[0], p[1], 'Y']; }));
  }

  var acc = ensure(SHEETS.ACCOUNT, ACCOUNT_HEADERS);
  var legacy = readLegacyAccounts_();
  if (acc.sheet.getLastRow() === 1) {
    var tmpNo = 0;
    var rows = legacy.map(function (a) {
      var code = a.code || ('TMP-' + ('000' + (++tmpNo)).slice(-4));
      return [code, a.name, a.owner, a.dcm ? 'DCM' : '방문처', '', a.focus, '', '', 'Y', a.extra || ''];
    });
    if (rows.length) {
      acc.sheet.getRange(2, 1, rows.length, 1).setNumberFormat('@');
      acc.sheet.getRange(2, ACCOUNT_HEADERS.length, rows.length, 1).setNumberFormat('@');
      acc.sheet.getRange(2, 1, rows.length, ACCOUNT_HEADERS.length).setValues(rows);
    }
    Logger.log('거래처 ' + rows.length + '곳 등록 (임시코드 ' + tmpNo + '곳, DCM ' + legacy.filter(function (a) { return a.dcm; }).length + '곳)');
  } else {
    // 이전 버전으로 이미 만든 거래처 시트: '구분' 열이 없으면 추가하고 DCM 표시
    var head = acc.sheet.getRange(1, 1, 1, acc.sheet.getLastColumn()).getValues()[0].map(String);
    if (head.indexOf('구분') === -1) {
      var col = head.length + 1;
      acc.sheet.getRange(1, col).setValue('구분').setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
      var dcmCodes = {};
      legacy.forEach(function (a) { if (a.dcm && a.code) dcmCodes[a.code] = true; });
      var cCode = head.indexOf('거래처코드');
      var n = acc.sheet.getLastRow() - 1;
      if (n > 0) {
        var codes = acc.sheet.getRange(2, cCode + 1, n, 1).getValues();
        acc.sheet.getRange(2, col, n, 1).setValues(codes.map(function (r) { return [dcmCodes[normalizeCode_(r[0])] ? 'DCM' : '방문처']; }));
      }
      Logger.log("거래처 시트에 '구분' 열을 추가했습니다. (MBO 열은 더 이상 쓰지 않으니 지워도 됩니다)");
    }
    head = acc.sheet.getRange(1, 1, 1, acc.sheet.getLastColumn()).getValues()[0].map(String);
    if (head.indexOf('추가코드') === -1) {
      var col2 = head.length + 1;
      acc.sheet.getRange(1, col2).setValue('추가코드').setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
      var iCode = head.indexOf('거래처코드');
      var n2 = acc.sheet.getLastRow() - 1;
      if (n2 > 0) {
        var codes2 = acc.sheet.getRange(2, iCode + 1, n2, 1).getValues();
        acc.sheet.getRange(2, col2, n2, 1).setNumberFormat('@')
          .setValues(codes2.map(function (r) { return [(LEGACY_EXTRA_CODES[normalizeCode_(r[0])] || []).join(', ')]; }));
      }
      Logger.log("거래처 시트에 '추가코드' 열을 추가했습니다.");
    }
  }

  var tz = ss.getSpreadsheetTimeZone(), stz = Session.getScriptTimeZone();
  Logger.log('설정 완료. 스프레드시트 시간대=' + tz + ', 스크립트 시간대=' + stz);
  if (tz !== stz) Logger.log('⚠ 두 시간대가 다릅니다. 둘 다 Asia/Seoul 로 맞춰 주세요. (날짜가 하루씩 밀릴 수 있음)');
}

/* =========================== 공통 유틸 =========================== */

function sheet_(name) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('"' + name + '" 시트가 없습니다. 관리자가 setupSheets() 를 먼저 실행해야 합니다.');
  return sh;
}

function readTable_(name) {
  var sh = sheet_(name);
  var vals = sh.getDataRange().getValues();
  var headers = (vals[0] || []).map(function (x) { return String(x).trim(); });
  var rows = [];
  for (var i = 1; i < vals.length; i++) {
    var o = { _row: i + 1 }, empty = true;
    for (var j = 0; j < headers.length; j++) {
      o[headers[j]] = vals[i][j];
      if (vals[i][j] !== '' && vals[i][j] !== null) empty = false;
    }
    if (!empty) rows.push(o);
  }
  return { sheet: sh, headers: headers, rows: rows };
}

function tz_() { return Session.getScriptTimeZone() || 'Asia/Seoul'; }

function fmtDate_(v) {
  if (v === null || v === undefined || v === '') return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return isNaN(v.getTime()) ? '' : Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  }
  var m = String(v).trim().match(/^(\d{4})[-.\/]\s*(\d{1,2})[-.\/]\s*(\d{1,2})/);
  return m ? m[1] + '-' + pad2_(m[2]) + '-' + pad2_(m[3]) : '';
}

function parseDate_(s) {
  var m = String(s || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])) : null;
}

function toNumber_(v) {
  if (typeof v === 'number') return v;
  var n = parseFloat(String(v || '').replace(/[,"\s원]/g, ''));
  return isNaN(n) ? 0 : n;
}

function isTmpCode_(c) { return /^TMP-\d+$/.test(String(c)); }

function nextTmpCode_(codes) {
  var max = 0;
  codes.forEach(function (c) { var m = /^TMP-(\d+)$/.exec(c); if (m) max = Math.max(max, Number(m[1])); });
  return 'TMP-' + ('000' + (max + 1)).slice(-4);
}

/** 시트 수식 주입 방지 + 길이 제한 */
function clean_(v) {
  var s = String(v === undefined || v === null ? '' : v).trim().slice(0, 3000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
