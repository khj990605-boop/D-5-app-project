/**
 * 뉴비키 조회수 (Apps Script 파일 views.gs)
 *
 * - 앱에서 용어 상세를 열면 그 용어의 조회수를 1 올려요.
 *   (같은 브라우저에서 같은 용어는 하루에 한 번만 세요. 그 판단은 앱이 해요)
 * - 누가 봤는지는 저장하지 않아요. 용어별 숫자만 「조회수」 탭에 있어요.
 *     조회수        : 지금까지 누적 → 카테고리 "많이 찾은 순" 정렬
 *     이번 주 조회수 : 월요일(한국 시간)마다 0부터 다시 → 홈 "🔥 이번 주 많이 찾은 말"
 * - 앱이 열릴 때 용어와 함께 조회수를 받아 가요.
 *
 * Code.gs에 연결하는 두 곳
 *   1) doPost 맨 위:  var viewRes = handleViewPost_(e); if (viewRes) return viewRes;
 *   2) doGet 마지막 return 줄에 views: viewCounts_() 추가
 */

var VIEW_SHEET = "조회수";
var VIEW_HEAD = ["용어ID", "조회수", "마지막 조회", "이번 주 시작", "이번 주 조회수"];
var VIEW_MAX_ROWS = 2000; // 이상한 id가 끝없이 쌓이지 않게 막는 안전장치

/* 앱 → 조회 1회 기록. type이 "view"가 아니면 null을 돌려줘서 다음 처리로 넘겨요 */
function handleViewPost_(e) {
  var body;
  try { body = JSON.parse((e && e.postData && e.postData.contents) || "{}"); } catch (err) { return null; }
  if (!body || body.type !== "view") return null;
  try {
    var term = String(body.term || "");
    if (!/^[a-z0-9]{1,40}$/.test(term)) return viewJson_({ ok: false, error: "term" });
    // 댓글 코드(comments.gs)의 용어 확인을 같이 써요. 없는 용어면 세지 않아요
    if (typeof cmtTermExists_ === "function" && !cmtTermExists_(term)) return viewJson_({ ok: false, error: "term" });

    var lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      var sh = viewSheet_();
      var last = sh.getLastRow();
      var ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues() : [];
      var row = -1;
      for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === term) { row = i + 2; break; }
      var week = viewWeekStart_();
      if (row > 0) {
        var r = sh.getRange(row, 1, 1, VIEW_HEAD.length).getValues()[0];
        var sameWeek = viewDay_(r[3]) === week;
        sh.getRange(row, 2, 1, 4).setValues([[
          (Number(r[1]) || 0) + 1,                       // 누적
          new Date(),                                     // 마지막 조회
          week,                                           // 이번 주 시작 (월요일)
          (sameWeek ? Number(r[4]) || 0 : 0) + 1          // 주가 바뀌었으면 0부터
        ]]);
      } else {
        if (last - 1 >= VIEW_MAX_ROWS) return viewJson_({ ok: false, error: "full" });
        sh.appendRow([term, 1, new Date(), week, 1]);
      }
    } finally {
      lock.releaseLock();
    }
    return viewJson_({ ok: true });
  } catch (err) {
    return viewJson_({ ok: false, error: "server" });
  }
}

/* 용어별 조회수 { id: { total: 누적, week: 이번 주 } }. 읽지 못하면 빈 값이라 앱은 0으로 보여줘요
   이번 주에 한 번도 안 본 용어는 지난주 숫자가 남아 있어도 week를 0으로 보내요 */
function viewCounts_() {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VIEW_SHEET);
    if (!sh || sh.getLastRow() < 2) return {};
    var week = viewWeekStart_();
    var out = {};
    sh.getRange(2, 1, sh.getLastRow() - 1, VIEW_HEAD.length).getValues().forEach(function (r) {
      var id = String(r[0] || ""), n = Number(r[1]) || 0;
      if (!id || n < 1) return;
      out[id] = { total: n, week: viewDay_(r[3]) === week ? Number(r[4]) || 0 : 0 };
    });
    return out;
  } catch (err) {
    return {};
  }
}

function viewSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(VIEW_SHEET);
  if (!sh) {
    sh = ss.insertSheet(VIEW_SHEET);
    sh.setFrozenRows(1);
  }
  // 머리글 (예전 3칸짜리 탭이면 '이번 주' 칸 두 개를 채워 넣어요)
  if (String(sh.getRange(1, VIEW_HEAD.length).getValue()) !== VIEW_HEAD[VIEW_HEAD.length - 1]) {
    sh.getRange(1, 1, 1, VIEW_HEAD.length).setValues([VIEW_HEAD]).setFontWeight("bold");
  }
  return sh;
}

// 이번 주 월요일 날짜 (한국 시간, "yyyy-MM-dd")
function viewWeekStart_() {
  var now = new Date();
  var dow = Number(Utilities.formatDate(now, "Asia/Seoul", "u")); // 1=월 … 7=일
  return Utilities.formatDate(new Date(now.getTime() - (dow - 1) * 86400000), "Asia/Seoul", "yyyy-MM-dd");
}

// 시트가 날짜 글자를 날짜로 바꿔 저장해도 "yyyy-MM-dd"로 맞춰 비교해요
function viewDay_(v) {
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) return Utilities.formatDate(v, "Asia/Seoul", "yyyy-MM-dd");
  return String(v || "").slice(0, 10);
}

function viewJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
