/**
 * 2026 정책위원회 워크숍 참석 회신 백엔드
 * 서울특별시사회복지사협회 정책위원회
 *
 * doGet  : 회신 전체를 JSON으로 반환
 * doPost : 회신 1건을 id 기준으로 저장(있으면 갱신, 없으면 추가)
 *
 * 배포: 배포 > 새 배포 > 유형 '웹 앱'
 *   - 실행 계정: 나
 *   - 액세스 권한: 모든 사용자
 *   배포 후 /exec 로 끝나는 URL을 index.html 의 API_URL 에 넣습니다.
 */

// 스크립트를 시트에 연결해 만들었으면 비워 두세요.
// 별도 스프레드시트를 쓰려면 시트 URL의 /d/ 와 /edit 사이 문자열을 넣습니다.
var SHEET_ID = '';

var SHEET_NAME = '응답';
var HEADERS = ['타임스탬프', 'id', '이름', '참석여부', '도착일', '교통편', '열차', '승차역', '열차출발', '도착역', '도착시각', '메모'];

var COL_ID = 2;   // B
var COL_DEP = 9;  // I 열차출발
var COL_ARR = 11; // K 도착시각

var ALLOWED = { full: 1, part: 1, tbd: 1, no: 1 };

/** 시트를 가져오고, 없으면 머리글까지 만들어 반환 */
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
    sh.setColumnWidth(7, 200);
    sh.setColumnWidth(12, 260);
  }

  // 시각 칸이 날짜로 바뀌지 않게 텍스트 서식 고정
  sh.getRange(1, COL_ARR, sh.getMaxRows(), 1).setNumberFormat('@');
  sh.getRange(1, COL_DEP, sh.getMaxRows(), 1).setNumberFormat('@');

  return sh;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** 셀 값이 Date로 읽히더라도 HH:MM 문자열로 통일 */
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

/** 회신 조회 */
function doGet() {
  try {
    var sh = getSheet_();
    var last = sh.getLastRow();
    var rows = [];

    if (last > 1) {
      var values = sh.getRange(2, 1, last - 1, HEADERS.length).getValues();
      for (var i = 0; i < values.length; i++) {
        var r = values[i];
        if (!r[COL_ID - 1]) continue;
        rows.push({
          id: Number(r[1]),
          name: asText_(r[2]),
          status: ALLOWED[asText_(r[3])] ? asText_(r[3]) : 'tbd',
          day: asText_(r[4]),
          transport: asText_(r[5]),
          train: asText_(r[6]),
          from: asText_(r[7]),
          dep: asTime_(r[8]),
          to: asText_(r[9]),
          arr: asTime_(r[10]),
          note: asText_(r[11]),
          at: r[0] instanceof Date ? r[0].toISOString() : asText_(r[0])
        });
      }
    }

    return json_({ ok: true, count: rows.length, rows: rows });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

/** 회신 저장 (id 기준 upsert) */
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
    if (!ALLOWED[d.status]) return json_({ ok: false, error: '참석여부 값이 올바르지 않습니다.' });

    var sh = getSheet_();
    var last = sh.getLastRow();
    var target = 0;

    if (last > 1) {
      var ids = sh.getRange(2, COL_ID, last - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) {
        if (Number(ids[i][0]) === id) { target = i + 2; break; }
      }
    }
    if (!target) target = (last < 1 ? 1 : last) + 1;

    var row = [
      new Date(),
      id,
      asText_(d.name),
      asText_(d.status),
      asText_(d.day),
      asText_(d.transport),
      asText_(d.trainLabel) || asText_(d.train),
      asText_(d.from),
      asTime_(d.dep),
      asText_(d.to),
      asTime_(d.arr),
      asText_(d.note)
    ];

    sh.getRange(target, 1, 1, HEADERS.length).setValues([row]);

    return json_({ ok: true, id: id, row: target });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

/**
 * 처음 한 번 실행해 시트와 머리글을 만듭니다.
 * 편집기에서 setup 을 선택하고 실행하면 됩니다.
 */
function setup() {
  var sh = getSheet_();
  Logger.log('준비 완료: ' + sh.getParent().getName() + ' / ' + sh.getName());
}

/** 저장·조회가 잘 되는지 편집기에서 확인하는 용도 */
function selfTest() {
  var res = doPost({ postData: { contents: JSON.stringify({
    id: 999, name: '테스트', status: 'part', day: '1일차', transport: '기차',
    train: '1283', trainLabel: '무궁화 1283 · 용산 16:25 → 광천 18:55',
    from: '용산역', dep: '16:25', to: '광천역', arr: '18:55', note: '자체 점검'
  }) } });
  Logger.log(res.getContent());
  Logger.log(doGet().getContent());
  Logger.log('점검이 끝나면 시트에서 id 999 행을 지우세요.');
}
