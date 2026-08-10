/**
 * 2026 정책위원회 워크숍 — 참석자 명단 연동
 * 서울특별시사회복지사협회 정책위원회
 *
 * 이 스크립트는 「참석자 명단」 스프레드시트를 그대로 읽고 씁니다.
 * 별도의 응답 시트를 만들지 않습니다. 명단 시트가 유일한 원본입니다.
 *
 *   doGet  : 명단을 읽어 JSON으로 반환
 *   doPost : 페이지에서 고친 내용을 명단 시트의 해당 행에 덮어씀
 *
 * 건드리는 열은 참석구분 · 이동방법 · 기차시간(하행/상행) · 숙소 · 티셔츠 여섯 개뿐입니다.
 * 연락처와 성별은 읽지도 쓰지도 않습니다.
 *
 * 배포: 배포 > 새 배포 > 웹 앱 / 실행 계정: 나 / 액세스 권한: 모든 사용자
 */

// 참석자 명단 스프레드시트 ID (주소의 /d/ 와 /edit 사이 문자열)
var SHEET_ID = '1H9hUU8555xnxUlGQ47j9QSLF--BYF4PqhDdjy_QDAP4';

/* 성명 → 페이지 id. 페이지의 BASE 와 같아야 합니다. */
var IDS = {
  '김아래미':1, '이상현':2, '김유리':3, '권민지':4, '송경태':5, '조상우':6,
  '고석우':7, '이재중':8, '이상표':9, '김민재':10, '황흥기':11, '정선영':12,
  '노혜진':13, '심휘선':14, '정순영':15, '오순희':16, '조소연':17
};

/* 하행 열차 출발시각 → 편성 id */
var DEP_TO_TRAIN = {
  '09:14':'g1', '09:27':'g2',
  '15:28':'1207', '16:25':'1283', '17:52':'1209', '18:21':'1285', '20:43':'1287'
};

/* 시트에 다시 적을 때 쓰는 문구 */
var STATUS_TXT = { full:'전일정', day1:'부분 참석', late:'후발대 참석', tbd:'확인 중', no:'불참' };
var DOWN_TXT = {
  g1:'용산 09:14→광천 11:36', g2:'영등포 09:27→광천 11:36',
  '1207':'용산 15:28→광천 17:50', '1283':'용산 16:25→광천 18:55',
  '1209':'용산 17:52→광천 20:13', '1285':'용산 18:21→광천 20:54',
  '1287':'용산 20:43→광천 23:11'
};
var UP_TXT = {
  u1:'광천 13:14 → 영등포 15:36', u2:'광천 13:14 → 용산 15:49',
  x1:'1일차만 참석', x2:'자체 이동', x3:'홍성 잔류'
};
var ROOM_TXT = { r7:'7인실', r1a:'1인실_1', r1b:'1인실_2', r1c:'1인실_3', rm:'모듈', none:'-' };
var ROOM_ID = { '7인실':'r7', '1인실_1':'r1a', '1인실_2':'r1b', '1인실_3':'r1c', '모듈':'rm' };

var ALLOWED = { full:1, day1:1, late:1, tbd:1, no:1 };

/* ───────────────────── 시트 찾기 ───────────────────── */

/** 머리글에 성명·참석구분이 있는 시트와 그 머리글 행을 찾는다 */
function roster_() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheets = ss.getSheets();

  for (var i = 0; i < sheets.length; i++) {
    var sh = sheets[i];
    var lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
    if (lastRow < 1 || lastCol < 1) continue;

    var scan = Math.min(lastRow, 12);
    var vals = sh.getRange(1, 1, scan, lastCol).getValues();
    for (var r = 0; r < vals.length; r++) {
      var head = [];
      for (var c = 0; c < vals[r].length; c++) head.push(String(vals[r][c]).trim());
      if (head.indexOf('성명') >= 0 && head.indexOf('참석구분') >= 0) {
        return { sheet: sh, headRow: r + 1, head: head, col: cols_(head), width: lastCol };
      }
    }
  }
  throw new Error('참석자 명단 시트를 찾지 못했습니다. 머리글에 성명·참석구분이 있어야 합니다.');
}

function cols_(head) {
  var c = {};
  for (var i = 0; i < head.length; i++) {
    var h = head[i];
    if (h === '성명') c.name = i + 1;
    else if (h === '참석구분') c.status = i + 1;
    else if (h === '이동방법') c.move = i + 1;
    else if (h === '숙소') c.room = i + 1;
    else if (h === '티셔츠') c.shirt = i + 1;
    else if (h.indexOf('기차시간') === 0) {
      if (!c.down) c.down = i + 1;
      else if (!c.up) c.up = i + 1;
    }
  }
  return c;
}

/* ───────────────────── 값 해석 ───────────────────── */

function txt_(v) {
  if (v instanceof Date) {
    return ('0' + v.getHours()).slice(-2) + ':' + ('0' + v.getMinutes()).slice(-2);
  }
  return v === null || v === undefined ? '' : String(v).trim();
}

/** 문자열에서 HH:MM 을 모두 뽑는다 */
function times_(v) {
  var out = [], re = /(\d{1,2}):(\d{2})/g, m;
  while ((m = re.exec(String(v)))) out.push(('0' + m[1]).slice(-2) + ':' + m[2]);
  return out;
}

function parseStatus_(v) {
  if (!v) return 'tbd';
  if (v.indexOf('전일정') >= 0) return 'full';
  if (v.indexOf('후발') >= 0) return 'late';
  if (v.indexOf('불참') >= 0) return 'no';
  if (v.indexOf('부분') >= 0 || v.indexOf('1일차') >= 0) return 'day1';
  return 'tbd';
}

function parseMove_(v) {
  if (!v) return '';
  if (v.indexOf('자차') >= 0) return '자차';
  if (v.indexOf('기차') >= 0) return '기차';
  return '기타';
}

/** 하행 칸 → { train, arr } */
function parseDown_(v) {
  var t = times_(v);
  if (v.indexOf('도착') >= 0) return { train: '', arr: t.length ? t[0] : '' };
  if (!t.length) return { train: '', arr: '' };
  return { train: DEP_TO_TRAIN[t[0]] || '', arr: '' };
}

/** 상행 칸 → 복귀 id */
function parseUp_(v) {
  if (!v) return '';
  if (v.indexOf('잔류') >= 0) return 'x3';
  if (v.indexOf('1일차만') >= 0) return 'x1';
  if (v.indexOf('자체') >= 0 || v.indexOf('자차') >= 0) return 'x2';
  if (v.indexOf('영등포') >= 0) return 'u1';
  if (v.indexOf('용산') >= 0) return 'u2';
  return '';
}

function parseRoom_(v) {
  var s = String(v).replace(/\s/g, '');
  if (ROOM_ID[s]) return ROOM_ID[s];
  if (!s || s === '-' || s === '–' || s === '—') return 'none';
  return '';
}

function parseShirt_(v) {
  var s = String(v).trim();
  return (s === '95' || s === '100' || s === '105') ? s : '';
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ───────────────────── 조회 ───────────────────── */

function doGet() {
  try {
    var R = roster_();
    var sh = R.sheet, col = R.col;
    var first = R.headRow + 1;
    var last = sh.getLastRow();
    var rows = [];

    if (last >= first && col.name) {
      var vals = sh.getRange(first, 1, last - first + 1, R.width).getValues();
      for (var i = 0; i < vals.length; i++) {
        var r = vals[i];
        var get = function (c) { return c ? txt_(r[c - 1]) : ''; };

        var name = get(col.name);
        if (!name || !IDS[name]) continue;

        var status = parseStatus_(get(col.status));
        var down = parseDown_(get(col.down));

        rows.push({
          id: IDS[name],
          name: name,
          status: status,
          transport: parseMove_(get(col.move)),
          train: down.train,
          arr: down.arr,
          back: parseUp_(get(col.up)),
          room: parseRoom_(get(col.room)),
          shirt: parseShirt_(get(col.shirt)),
          note: ''
        });
      }
    }

    return json_({ ok: true, count: rows.length, rows: rows });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/* ───────────────────── 저장 ───────────────────── */

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    if (!e || !e.postData || !e.postData.contents) {
      return json_({ ok: false, error: '전달된 내용이 없습니다.' });
    }

    var d = JSON.parse(e.postData.contents);
    var name = txt_(d.name);
    if (!name) return json_({ ok: false, error: '성명이 없습니다.' });
    if (!ALLOWED[d.status]) return json_({ ok: false, error: '참석구분 값이 올바르지 않습니다: ' + d.status });

    var R = roster_();
    var sh = R.sheet, col = R.col;
    var first = R.headRow + 1;
    var last = sh.getLastRow();

    // 성명으로 행 찾기
    var target = 0;
    if (last >= first && col.name) {
      var names = sh.getRange(first, col.name, last - first + 1, 1).getValues();
      for (var i = 0; i < names.length; i++) {
        if (txt_(names[i][0]) === name) { target = first + i; break; }
      }
    }
    if (!target) {
      target = last + 1;
      if (col.name) sh.getRange(target, col.name).setValue(name);
    }

    var put = function (c, val) {
      if (!c) return;
      sh.getRange(target, c).setNumberFormat('@').setValue(val);
    };

    var joins = (d.status === 'full' || d.status === 'day1' || d.status === 'late');

    put(col.status, STATUS_TXT[d.status] || '');

    if (!joins) {
      put(col.move, '');
      put(col.down, '');
      put(col.up, '');
      put(col.room, '');
      put(col.shirt, '');
    } else {
      var move = d.transport === '자차' ? '자차'
               : d.transport === '기차' ? (d.train === 'g1' || d.train === 'g2' ? '기차(단체)' : '기차(개별)')
               : txt_(d.transport);
      put(col.move, move);

      if (d.transport === '자차' || d.transport === '기타') {
        put(col.down, txt_(d.arr) ? txt_(d.arr) + ' 도착 예정' : '');
      } else {
        put(col.down, DOWN_TXT[d.train] || '');
      }

      put(col.up, UP_TXT[d.back] || '');
      put(col.room, ROOM_TXT[d.room] !== undefined ? ROOM_TXT[d.room] : '');
      put(col.shirt, txt_(d.shirt) || '미정');
    }

    SpreadsheetApp.flush();
    return json_({ ok: true, name: name, row: target });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

/* ───────────────────── 점검 ───────────────────── */

/** 명단 시트를 제대로 찾는지, 열이 다 잡히는지 확인 */
function setup() {
  var R = roster_();
  Logger.log('스프레드시트: ' + R.sheet.getParent().getName());
  Logger.log('시트 탭: ' + R.sheet.getName() + ' / 머리글 행: ' + R.headRow);
  var need = ['name', 'status', 'move', 'down', 'up', 'room', 'shirt'];
  var miss = [];
  for (var i = 0; i < need.length; i++) if (!R.col[need[i]]) miss.push(need[i]);
  Logger.log('열 인식: ' + (miss.length ? '누락 ' + miss.join(', ') : '정상 (전부 찾음)'));
  Logger.log(JSON.stringify(R.col));
}

/** 명단을 읽어서 페이지에 넘길 값이 제대로 나오는지 확인 */
function preview() {
  var res = JSON.parse(doGet().getContent());
  if (!res.ok) { Logger.log('실패: ' + res.error); return; }
  Logger.log('읽은 인원: ' + res.count + '명');
  for (var i = 0; i < res.rows.length; i++) {
    var r = res.rows[i];
    Logger.log([r.id, r.name, r.status, r.transport, r.train || r.arr, r.back, r.room, r.shirt].join(' | '));
  }
}
