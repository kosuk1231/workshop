/**
 * 2026 정책위원회 워크숍 참석 회신 백엔드
 * 서울특별시사회복지사협회 정책위원회
 *
 * doGet  : 회신 전체를 JSON으로 반환
 * doPost : 회신 1건을 id 기준으로 저장(있으면 갱신, 없으면 추가)
 *
 * 열 위치를 번호로 찾지 않고 머리글 이름으로 찾습니다.
 * 열 순서를 바꾸거나 예전 시트를 그대로 써도 값이 밀리지 않습니다.
 * 없는 머리글은 실행할 때 자동으로 뒤에 추가됩니다.
 *
 * 배포: 배포 > 새 배포 > 유형 '웹 앱'
 *   - 실행 계정: 나
 *   - 액세스 권한: 모든 사용자
 *   배포 후 /exec 로 끝나는 URL을 index.html 의 API_URL 에 넣습니다.
 */

// 스크립트를 시트에 연결해 만들었으면 비워 두세요.
var SHEET_ID = '';

var SHEET_NAME = '응답';

/* 머리글 이름 = 이 목록이 기준입니다. 순서를 바꿔도 동작합니다. */
var H = {
  at:         '타임스탬프',
  id:         'id',
  name:       '이름',
  status:     '참석여부',
  day:        '도착일',
  transport:  '교통편',
  trainLabel: '열차',
  train:      '열차ID',
  from:       '승차역',
  dep:        '열차출발',
  to:         '도착역',
  arr:        '도착시각',
  shirt:      '티셔츠',
  note:       '메모'
};

var HEADERS = [
  H.at, H.id, H.name, H.status, H.day, H.transport,
  H.trainLabel, H.train, H.from, H.dep, H.to, H.arr, H.shirt, H.note
];

/* 시각처럼 보이는 값이 날짜로 바뀌지 않게 텍스트 서식으로 둘 열 */
var TEXT_COLS = [H.dep, H.arr, H.train];

var ALLOWED = { full: 1, part: 1, tbd: 1, no: 1 };

/* ─────────────────────────── 시트 ─────────────────────────── */

function getSheet_() {
  var ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('스프레드시트를 찾을 수 없습니다. SHEET_ID를 확인하세요.');

  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 150);
    sh.setColumnWidth(3, 90);
  }

  migrate_(sh);
  return sh;
}

/** 머리글 이름 → 열 번호 */
function headerMap_(sh) {
  var width = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, width).getValues()[0];
  var map = {};
  for (var i = 0; i < hdr.length; i++) {
    var k = String(hdr[i]).trim();
    if (k) map[k] = i + 1;
  }
  return map;
}

/** 빠진 머리글을 뒤에 추가하고 텍스트 서식을 걸어 둔다 */
function migrate_(sh) {
  var map = headerMap_(sh);
  var add = [];
  for (var i = 0; i < HEADERS.length; i++) {
    if (!map[HEADERS[i]]) add.push(HEADERS[i]);
  }
  if (add.length) {
    var start = sh.getLastColumn() + 1;
    var need = start + add.length - 1;
    if (sh.getMaxColumns() < need) sh.insertColumnsAfter(sh.getMaxColumns(), need - sh.getMaxColumns());
    sh.getRange(1, start, 1, add.length).setValues([add]).setFontWeight('bold');
    map = headerMap_(sh);
  }
  for (var j = 0; j < TEXT_COLS.length; j++) {
    var c = map[TEXT_COLS[j]];
    if (c) sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@');
  }
  return map;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function asTime_(v) {
  if (v instanceof Date) {
    var h = ('0' + v.getHours()).slice(-2);
    var m = ('0' + v.getMinutes()).slice(-2);
    return h + ':' + m;
  }
  return v === null || v === undefined ? '' : String(v).trim();
}

function asText_(v) {
  return v === null || v === undefined ? '' : String(v).trim();
}

/* ─────────────────────────── 조회 ─────────────────────────── */

function doGet() {
  try {
    var sh = getSheet_();
    var map = headerMap_(sh);
    var last = sh.getLastRow();
    var rows = [];

    if (last > 1) {
      var width = sh.getLastColumn();
      var values = sh.getRange(2, 1, last - 1, width).getValues();

      for (var i = 0; i < values.length; i++) {
        var r = values[i];
        var get = function (key) {
          var c = map[key];
          return c ? r[c - 1] : '';
        };
        var id = Number(get(H.id));
        if (!id) continue;
        var st = asText_(get(H.status));
        var at = get(H.at);
        rows.push({
          id: id,
          name: asText_(get(H.name)),
          status: ALLOWED[st] ? st : 'tbd',
          day: asText_(get(H.day)),
          transport: asText_(get(H.transport)),
          train: asText_(get(H.train)),
          trainLabel: asText_(get(H.trainLabel)),
          from: asText_(get(H.from)),
          dep: asTime_(get(H.dep)),
          to: asText_(get(H.to)),
          arr: asTime_(get(H.arr)),
          shirt: asText_(get(H.shirt)),
          note: asText_(get(H.note)),
          at: at instanceof Date ? at.toISOString() : asText_(at)
        });
      }
    }

    return json_({ ok: true, count: rows.length, rows: rows });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/* ─────────────────────────── 저장 ─────────────────────────── */

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    if (!e || !e.postData || !e.postData.contents) {
      return json_({ ok: false, error: '전달된 내용이 없습니다.' });
    }

    var d = JSON.parse(e.postData.contents);
    var id = Number(d.id);
    if (!id) return json_({ ok: false, error: 'id가 없습니다.' });
    if (!ALLOWED[d.status]) return json_({ ok: false, error: '참석여부 값이 올바르지 않습니다: ' + d.status });

    var sh = getSheet_();
    var map = headerMap_(sh);
    var width = sh.getLastColumn();

    var idCol = map[H.id];
    var last = sh.getLastRow();
    var target = 0;
    if (idCol && last > 1) {
      var ids = sh.getRange(2, idCol, last - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (Number(ids[i][0]) === id) { target = i + 2; break; }
      }
    }
    var isNew = !target;
    if (isNew) target = last + 1;

    var row = isNew ? [] : sh.getRange(target, 1, 1, width).getValues()[0];
    for (var k = 0; k < width; k++) if (row[k] === undefined || row[k] === null) row[k] = '';

    var put = function (key, val) {
      var c = map[key];
      if (c) row[c - 1] = val;
    };

    put(H.at, new Date());
    put(H.id, id);
    put(H.name, asText_(d.name));
    put(H.status, asText_(d.status));
    put(H.day, asText_(d.day));
    put(H.transport, asText_(d.transport));
    put(H.train, asText_(d.train));
    put(H.trainLabel, asText_(d.trainLabel));
    put(H.from, asText_(d.from));
    put(H.dep, asTime_(d.dep));
    put(H.to, asText_(d.to));
    put(H.arr, asTime_(d.arr));
    put(H.shirt, asText_(d.shirt));
    put(H.note, asText_(d.note));

    sh.getRange(target, 1, 1, width).setValues([row]);
    SpreadsheetApp.flush();

    return json_({ ok: true, id: id, row: target, isNew: isNew });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

/* ─────────────────────────── 점검용 ─────────────────────────── */

/** 처음 한 번, 또는 코드를 새로 붙여넣은 뒤 실행하세요. */
function setup() {
  var sh = getSheet_();
  var map = headerMap_(sh);
  var missing = [];
  for (var i = 0; i < HEADERS.length; i++) if (!map[HEADERS[i]]) missing.push(HEADERS[i]);
  Logger.log('시트: ' + sh.getParent().getName() + ' / ' + sh.getName());
  Logger.log('머리글: ' + (missing.length ? '누락 ' + missing.join(', ') : '정상'));
}

/**
 * 회신을 전부 지우고 처음 상태로 되돌립니다. (머리글은 남습니다)
 * 테스트 데이터가 섞여 값이 꼬였을 때 한 번 실행하세요.
 */
function reset() {
  var sh = getSheet_();
  var last = sh.getLastRow();
  if (last > 1) {
    sh.getRange(2, 1, last - 1, sh.getLastColumn()).clearContent();
    Logger.log('회신 ' + (last - 1) + '건을 지웠습니다.');
  } else {
    Logger.log('지울 회신이 없습니다.');
  }
  Logger.log('페이지에서 새로고침을 누르면 명단이 처음 상태로 돌아옵니다.');
}

/** 특정 위원의 회신 한 건만 지웁니다. removeOne(12) 처럼 id를 넣어 실행하세요. */
function removeOne(id) {
  var sh = getSheet_();
  var map = headerMap_(sh);
  var last = sh.getLastRow();
  if (!map[H.id] || last < 2) { Logger.log('회신이 없습니다.'); return; }
  var ids = sh.getRange(2, map[H.id], last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (Number(ids[i][0]) === Number(id)) {
      sh.deleteRow(i + 2);
      Logger.log('id ' + id + ' 회신을 지웠습니다.');
      return;
    }
  }
  Logger.log('id ' + id + ' 회신을 찾지 못했습니다.');
}

/** 저장한 값이 그대로 돌아오는지 확인 */
function selfTest() {
  var sent = {
    id: 999, name: '테스트', status: 'part', day: '1일차', transport: '기차',
    train: '1283', trainLabel: '무궁화 1283 · 용산 16:25 → 광천 18:55',
    from: '용산역', dep: '16:25', to: '광천역', arr: '18:55', shirt: '100', note: '자체 점검'
  };
  Logger.log('저장: ' + doPost({ postData: { contents: JSON.stringify(sent) } }).getContent());

  var back = JSON.parse(doGet().getContent());
  var hit = null;
  for (var i = 0; i < back.rows.length; i++) if (back.rows[i].id === 999) hit = back.rows[i];
  if (!hit) { Logger.log('실패: 999번을 다시 읽지 못했습니다.'); return; }

  var keys = ['status', 'day', 'transport', 'train', 'from', 'dep', 'to', 'arr', 'shirt', 'note'];
  var bad = [];
  for (var j = 0; j < keys.length; j++) {
    if (String(hit[keys[j]]) !== String(sent[keys[j]])) {
      bad.push(keys[j] + ': 보냄 "' + sent[keys[j]] + '" / 읽음 "' + hit[keys[j]] + '"');
    }
  }
  Logger.log(bad.length ? '값이 다릅니다\n' + bad.join('\n') : '왕복 확인 정상 — 값이 그대로 돌아왔습니다.');
  Logger.log('점검이 끝나면 시트에서 id 999 행을 지우세요.');
}
