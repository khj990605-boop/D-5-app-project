/**
 * 뉴비키 조회수 (Apps Script 파일 views.gs)
 *
 * - 앱에서 용어 상세를 열면 그 용어의 조회수를 1 올려요.
 *   (같은 브라우저에서 같은 용어는 하루에 한 번만 세요. 그 판단은 앱이 해요)
 * - 누가 봤는지는 저장하지 않아요. 용어별 누적 숫자만 「조회수」 탭에 있어요.
 * - 앱이 열릴 때 용어와 함께 조회수를 받아 가요 → "많이 찾은 순" 정렬에 쓰여요.
 *
 * Code.gs에 연결하는 두 곳
 *   1) doPost 맨 위:  var viewRes = handleViewPost_(e); if (viewRes) return viewRes;
 *   2) doGet 마지막 return 줄에 views: viewCounts_() 추가
 */

var VIEW_SHEET = "조회수";
var VIEW_HEAD = ["용어ID", "조회수", "마지막 조회"];
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
      if (row > 0) {
        var cell = sh.getRange(row, 2);
        cell.setValue((Number(cell.getValue()) || 0) + 1);
        sh.getRange(row, 3).setValue(new Date());
      } else {
        if (last - 1 >= VIEW_MAX_ROWS) return viewJson_({ ok: false, error: "full" });
        sh.appendRow([term, 1, new Date()]);
      }
    } finally {
      lock.releaseLock();
    }
    return viewJson_({ ok: true });
  } catch (err) {
    return viewJson_({ ok: false, error: "server" });
  }
}

/* 용어별 조회수 { id: 숫자 }. 읽지 못하면 빈 값이라 앱은 0으로 보여줘요 */
function viewCounts_() {
  try {
    var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(VIEW_SHEET);
    if (!sh || sh.getLastRow() < 2) return {};
    var out = {};
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
      var id = String(r[0] || ""), n = Number(r[1]) || 0;
      if (id && n > 0) out[id] = n;
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
    sh.getRange(1, 1, 1, VIEW_HEAD.length).setValues([VIEW_HEAD]).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

function viewJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
