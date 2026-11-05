/**
 * ============================================================
 *  기존 "개인별 주차 시트" → "방문기록" 시트 1회 이관
 * ============================================================
 *  사용법 (Apps Script 편집기에서 직접 실행)
 *   1) setupSheets()     : 새 시트(방문기록/거래처/품목/담당자) 생성 + 기본값 채우기
 *   2) previewMigration(): 이관 결과를 미리 확인 (시트에 쓰지 않음, 실행 로그로만 표시)
 *   3) migrateLegacy()   : 실제로 방문기록 시트에 기록
 *
 *  기존 개인 시트는 수정하지 않습니다. (읽기만 함)
 */

/** 이관 대상 개인 시트 이름 (= 현재 담당자) */
var LEGACY_SHEETS = ['이수빈', '김준영', '유승태', '장문수', '정재은', '고현민'];

/** 기존 시트가 기록된 연도 */
var LEGACY_YEAR = 2026;

/**
 * 시트 주인과 실제 방문자가 달랐던 기간.
 * to: 이 날짜(또는 달)까지는 name 이 방문한 것으로 기록합니다. ('2026-04' = 4월 말까지, '2026-09-14' = 9/14까지)
 *  - 고현민 시트: 박준수(~4월) → 계상현 소장(5월) → 고현민(6월~)
 *  - 이수빈 시트: 구원모(~9/14, 인수인계 방문 포함) → 이수빈(9/15 "인수인계로 첫 방문"부터)
 */
var LEGACY_AUTHORS = {
  '고현민': [
    { to: '2026-04', name: '박준수' },
    { to: '2026-05', name: '계상현' }
  ],
  '이수빈': [
    { to: '2026-09-14', name: '구원모' }
  ]
};

/**
 * 기존 시트에 거래처코드가 비어 있던 곳 중, 평가실적의 거래처와 같은 곳으로 확인된 곳 (2026-11 소장님 확인)
 * key: '담당자|기존 시트의 거래처명'
 */
var LEGACY_CODE_FIX = {
  '장문수|거제의원': '10115243',
  '장문수|문영일이비인후과': '10094138',
  '정재은|안내과의원': '10030490',
  '이수빈|연아통증의학과의원': '10064568',
  '유승태|김양수소아과': '10057615',
  '고현민|한빛재활의학과': '10090186'
};

/** 한 거래처가 평가실적에 코드 여러 개로 잡히는 경우 (주 코드: 추가 코드들) */
var LEGACY_EXTRA_CODES = {
  '10053211': ['10089780']   // 강앤심내과의원
};

/* ------------------------------------------------------------------ */

function previewMigration() {
  var result = buildLegacyVisits_();
  Logger.log('이관 예정 방문 건수: ' + result.visits.length);
  Logger.log('날짜를 추정한 건수(본문에 일자 없음): ' + result.stats.estimatedDate);
  Logger.log('할일만 있고 결과가 없어 제외한 칸: ' + result.stats.todoOnly);
  Logger.log('"DCM처/휴가" 등 상태 표시라 제외한 칸: ' + result.stats.statusOnly);
  Object.keys(result.stats.bySheet).forEach(function (k) {
    Logger.log('  - ' + k + ': ' + result.stats.bySheet[k] + '건');
  });
  result.visits.slice(0, 15).forEach(function (v) {
    Logger.log(v.date + ' | ' + v.visitor + ' | ' + v.accountName + ' | ' + v.products + ' | ' + String(v.result).slice(0, 50));
  });
}

function migrateLegacy() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEETS.VISIT);
  if (!sh) throw new Error('먼저 setupSheets() 를 실행하세요.');

  // 중복 이관 방지
  var existing = sh.getLastRow() > 1 ? sh.getRange(2, VISIT_HEADERS.indexOf('출처') + 1, sh.getLastRow() - 1, 1).getValues() : [];
  var already = existing.some(function (r) { return String(r[0]).indexOf('이관') === 0; });
  if (already) {
    throw new Error('이미 이관된 기록이 있습니다. 다시 이관하려면 방문기록 시트에서 출처가 "이관"으로 시작하는 행을 지운 뒤 실행하세요.');
  }

  var result = buildLegacyVisits_();
  var now = new Date();
  var rows = result.visits.map(function (v, i) {
    var p = v.date.split('-');
    return [
      'L' + LEGACY_YEAR + '-' + ('0000' + (i + 1)).slice(-5),
      new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2])),
      v.visitor, v.accountCode, v.accountName, v.products, v.todo, '', v.result,
      '', '', v.companions, now, v.source, v.note
    ];
  });
  if (rows.length) {
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, VISIT_HEADERS.length).setValues(rows);
    sh.getRange(2, 2, sh.getLastRow() - 1, 1).setNumberFormat('yyyy-mm-dd');
  }
  Logger.log('이관 완료: ' + rows.length + '건');
}

/* ------------------------------------------------------------------ */

function buildLegacyVisits_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var products = loadProductMatchers_();
  var all = [];
  var stats = { estimatedDate: 0, todoOnly: 0, statusOnly: 0, bySheet: {} };
  // 거래처코드가 없던 곳은 거래처 시트의 임시코드(TMP-xxxx)로 연결
  var codeLookup = {};
  var acc = ss.getSheetByName(SHEETS.ACCOUNT);
  if (acc && acc.getLastRow() > 1) {
    var av = acc.getDataRange().getValues();
    var ah = av[0].map(String);
    var iC = ah.indexOf('거래처코드'), iN = ah.indexOf('거래처명'), iO = ah.indexOf('담당자');
    for (var i = 1; i < av.length; i++) {
      var c = normalizeCode_(av[i][iC]);
      if (isTmpCode_(c)) codeLookup[String(av[i][iO]).trim() + '|' + String(av[i][iN]).trim()] = c;
    }
  }
  LEGACY_SHEETS.forEach(function (name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    var values = sh.getDataRange().getValues();
    var out = parseLegacySheet_(name, values, LEGACY_YEAR, products, LEGACY_AUTHORS[name] || [], codeLookup);
    stats.estimatedDate += out.estimatedDate;
    stats.todoOnly += out.todoOnly;
    stats.statusOnly += out.statusOnly;
    stats.bySheet[name] = out.visits.length;
    all = all.concat(out.visits);
  });
  all.sort(function (a, b) { return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
  return { visits: all, stats: stats };
}

/**
 * 개인 시트 1장을 방문 단위로 분해 (순수 함수 — 시트 접근 없음)
 * values: getDataRange().getValues() 결과
 */
function parseLegacySheet_(owner, values, year, products, authors, codeLookup) {
  var hdrIdx = -1;
  for (var i = 0; i < Math.min(values.length, 10); i++) {
    if (values[i].map(String).indexOf('거래처명') !== -1) { hdrIdx = i; break; }
  }
  if (hdrIdx < 1) return { visits: [], estimatedDate: 0, todoOnly: 0 };
  var hdr = values[hdrIdx].map(function (x) { return String(x).trim(); });
  var weekRow = values[hdrIdx - 1];
  var colName = hdr.indexOf('거래처명');
  var colCode = hdr.indexOf('거래처 코드') !== -1 ? hdr.indexOf('거래처 코드') : hdr.indexOf('거래처코드');

  // 1) 주차 묶음 찾기: "결과" 열마다, 바로 앞의 "할일/한일" 열과 그 주차 제목
  var groups = [];
  for (var c = 0; c < hdr.length; c++) {
    if (hdr[c] !== '결과') continue;
    var todoCol = -1;
    for (var k = c - 1; k >= 0 && k >= c - 3; k--) {
      if (hdr[k] === '할일' || hdr[k] === '한일') { todoCol = k; break; }
    }
    var label = '';
    for (var k2 = c; k2 >= 0; k2--) {
      var w = String(weekRow[k2] || '').trim();
      if (w.indexOf('주차') !== -1) { label = w; break; }
    }
    var wk = parseWeekLabel_(label);
    if (!wk) continue;
    groups.push({ resultCol: c, todoCol: todoCol, label: label, month: wk.month, week: wk.week });
  }

  var visits = [];
  var estimated = 0, todoOnly = 0, statusOnly = 0;
  for (var r = hdrIdx + 1; r < values.length; r++) {
    var row = values[r];
    var accName = String(row[colName] || '').trim();
    var accCode = normalizeCode_(row[colCode]) || LEGACY_CODE_FIX[owner + '|' + accName] || (codeLookup || {})[owner + '|' + accName] || '';
    if (!accName || !accCode) continue;

    groups.forEach(function (g) {
      var resText = String(row[g.resultCol] || '').trim();
      var todoText = g.todoCol >= 0 ? String(row[g.todoCol] || '').trim() : '';
      if (!resText || resText === '-') { if (todoText) todoOnly++; return; }
      if (isStatusOnly_(resText)) { statusOnly++; return; }

      var resEntries = splitEntries_(resText);
      var todoEntries = todoText ? splitEntries_(todoText) : [];

      resEntries.forEach(function (re) {
        var todo = '';
        if (resEntries.length === 1 && todoEntries.length === 1) {
          todo = todoEntries[0].text;
        } else if (re.day) {
          var hit = todoEntries.filter(function (t) { return t.day === re.day; });
          todo = hit.map(function (t) { return t.text; }).join(' / ');
        }
        var d = resolveDate_(year, g.month, g.week, re.day);
        if (d.estimated) estimated++;
        var visitor = owner;
        for (var a = 0; a < authors.length; a++) {
          var to = authors[a].to;
          if (d.date.slice(0, to.length) <= to) { visitor = authors[a].name; break; }
        }
        var full = todo + ' ' + re.text;
        visits.push({
          date: d.date,
          visitor: visitor,
          accountCode: accCode,
          accountName: accName,
          products: matchProducts_(full, products).join(', '),
          todo: todo,
          result: re.text,
          companions: detectCompanions_(full).join(', '),
          source: '이관(' + owner + ' 시트 ' + g.label + ')',
          note: d.estimated ? '날짜추정' : ''
        });
      });
    });
  }
  return { visits: visits, estimatedDate: estimated, todoOnly: todoOnly, statusOnly: statusOnly };
}

/** "DCM처", "금주 휴가" 처럼 방문이 아닌 상태 표시 칸 */
function isStatusOnly_(text) {
  var t = String(text).trim();
  return t.length <= 20 && /(DC?M|휴가|공석|폐업)/i.test(t);
}

/** "10월2주차", "2월 3,4주차" → {month:10, week:2} */
function parseWeekLabel_(label) {
  var m = String(label).match(/(\d{1,2})\s*월\s*(\d)/);
  if (!m) return null;
  return { month: Number(m[1]), week: Number(m[2]) };
}

/**
 * 한 칸의 글을 방문 단위로 자르기.
 * 새 방문의 시작: "29일 :", "6일.", "26일;", "12:" 처럼 줄 맨 앞에 일자가 오는 줄
 */
function splitEntries_(text) {
  // 한 줄 안에 "6일 : … , 8일 : …" 처럼 이어 쓴 경우 줄바꿈으로 분리 (단, "5월 9일:" 은 제외)
  var s = String(text).replace(/(?<!월\s*)(?<!\d)(?<=.)(\d{1,2})\s*일\s*([:;])/g, '\n$1일$2');
  var lines = s.split(/\r?\n/);
  var entries = [];
  var re = /^\s*(\d{1,2})\s*(?:일\s*[.:;)\-]?|\s*[:;])\s*/;
  lines.forEach(function (line) {
    if (!line.trim()) return;
    var m = line.match(re);
    var day = m ? Number(m[1]) : 0;
    if (m && day >= 1 && day <= 31) {
      entries.push({ day: day, text: line.slice(m[0].length).trim() });
    } else if (entries.length) {
      entries[entries.length - 1].text += '\n' + line.trim();
    } else {
      entries.push({ day: 0, text: line.trim() });
    }
  });
  entries.forEach(function (e) { e.text = e.text.replace(/[\s,\/]+$/, ''); });
  return entries.filter(function (e) { return e.text; });
}

/** 주차 제목의 월 + 본문의 일자 → yyyy-MM-dd (월 경계 보정) */
function resolveDate_(year, month, week, day) {
  var estimated = false;
  var y = year, m = month;
  if (!day) {
    estimated = true;
    day = Math.min(1 + (week - 1) * 7, 28);
  } else if (week === 1 && day >= 22) {
    m -= 1;                                   // 1주차인데 25일 → 전월
  } else if (week >= 4 && day <= 10) {
    m += 1;                                   // 5주차인데 2일 → 다음 달
  }
  if (m < 1) { m = 12; y -= 1; }
  if (m > 12) { m = 1; y += 1; }
  var last = new Date(y, m, 0).getDate();
  if (day > last) day = last;
  return { date: y + '-' + pad2_(m) + '-' + pad2_(day), estimated: estimated };
}

function matchProducts_(text, products) {
  var t = String(text).toLowerCase();
  var hits = [];
  products.forEach(function (p) {
    var ok = p.keys.some(function (k) { return k && t.indexOf(k) !== -1; });
    if (ok) hits.push(p.name);
  });
  return hits;
}

function detectCompanions_(text) {
  var t = String(text);
  var withWord = /(동행|동도)/.test(t);
  var out = [];
  if (withWord && /CM/i.test(t)) out.push('CM');
  if (withWord && /PM/.test(t)) out.push('PM');
  if (withWord && /소장/.test(t)) out.push('소장');
  return out;
}

function loadProductMatchers_() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.PRODUCT);
  var list = [];
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(function (r) {
      if (!r[0]) return;
      list.push({ name: String(r[0]).trim(), aliases: String(r[1] || '') });
    });
  } else {
    list = DEFAULT_PRODUCTS.map(function (p) { return { name: p[0], aliases: p[1] }; });
  }
  return list.map(function (p) {
    var keys = [p.name].concat(p.aliases.split(',')).map(function (k) { return k.trim().toLowerCase(); }).filter(String);
    return { name: p.name, keys: keys };
  });
}

function normalizeCode_(v) {
  if (v === null || v === undefined || v === '') return '';
  var s = String(v).trim().replace(/\.0+$/, '');
  if (/^\d+$/.test(s)) return s;
  return /^TMP-\d+$/i.test(s) ? s.toUpperCase() : '';
}

/**
 * 기존 개인 시트의 거래처 목록 (setupSheets 가 '거래처' 시트를 채울 때 사용)
 * dcm: "DCM처" 같은 상태 표시가 3번 이상 적힌 거래처
 */
function readLegacyAccounts_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = [], seen = {};
  LEGACY_SHEETS.forEach(function (owner) {
    var sh = ss.getSheetByName(owner);
    if (!sh) return;
    var v = sh.getDataRange().getValues();
    var hi = -1;
    for (var i = 0; i < Math.min(v.length, 10); i++) if (v[i].map(String).indexOf('거래처명') !== -1) { hi = i; break; }
    if (hi < 0) return;
    var h = v[hi].map(function (x) { return String(x).trim(); });
    var cN = h.indexOf('거래처명'), cF = h.indexOf('집중품목');
    var cC = h.indexOf('거래처 코드') !== -1 ? h.indexOf('거래처 코드') : h.indexOf('거래처코드');
    for (var r = hi + 1; r < v.length; r++) {
      var name = String(v[r][cN] || '').trim();
      if (!name) continue;
      var code = normalizeCode_(v[r][cC]) || LEGACY_CODE_FIX[owner + '|' + name] || '';
      var key = code || (owner + '|' + name);
      if (seen[key]) continue;
      seen[key] = true;
      var dcmCells = 0;
      for (var c = 0; c < h.length; c++) {
        if (h[c] === '결과' && isStatusOnly_(v[r][c]) && /DC?M/i.test(String(v[r][c]))) dcmCells++;
      }
      out.push({ code: code, name: name, owner: owner, focus: cF >= 0 ? v[r][cF] : '', dcm: dcmCells >= 3,
        extra: (LEGACY_EXTRA_CODES[code] || []).join(', ') });
    }
  });
  return out;
}

/**
 * 9월 평가실적의 황내과의원 중복 행(1,428,746원) 삭제 — 1회 실행용.
 * 거래처코드·담당자·금액이 모두 일치하는 행만 지웁니다.
 */
function removeHwangDuplicate() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('9월 평가실적');
  if (!sh) throw new Error('9월 평가실적 탭이 없습니다.');
  var v = sh.getDataRange().getValues();
  for (var r = v.length - 1; r >= 1; r--) {
    if (normalizeCode_(v[r][3]) === '10008184' && String(v[r][0]).trim() === '정재은' && Math.round(Number(v[r][7])) === 1428746) {
      sh.deleteRow(r + 1);
      Logger.log((r + 1) + '행(황내과의원 1,428,746원)을 삭제했습니다.');
      return;
    }
  }
  Logger.log('삭제할 행을 찾지 못했습니다. (이미 지워졌을 수 있음)');
}

function pad2_(n) { return ('0' + Number(n)).slice(-2); }
