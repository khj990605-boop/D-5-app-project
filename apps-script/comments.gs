/**
 * 뉴비키 익명 댓글 · 좋아요 (Apps Script)
 *
 * 이 파일은 기존 Apps Script 프로젝트에 "새 파일"로 추가해요. 기존 코드는 지우거나 고치지 않아요.
 * 기존 doGet / doPost 맨 앞에 두 줄씩만 넣으면 돼요. (자세한 순서는 저장소 apps-script/README.md)
 *
 *   function doGet(e) {
 *     var cmtRes = handleCommentsGet_(e); if (cmtRes) return cmtRes;   // ← 추가
 *     ... 기존 코드 그대로 ...
 *   }
 *   function doPost(e) {
 *     var cmtRes = handleCommentPost_(e); if (cmtRes) return cmtRes;   // ← 추가
 *     ... 기존 코드 그대로 ...
 *   }
 *
 * - 댓글이 아닌 요청(용어 불러오기, 등록 요청, 나도 궁금해요)이면 null을 돌려줘서 기존 코드가 그대로 처리해요.
 * - 「댓글」·「좋아요」 탭은 처음 실행될 때 자동으로 만들어져요.
 * - 익명 번호와 좋아요 1회 제한은 LockService(스크립트 잠금)로 한 번에 하나씩 처리해서 겹치지 않아요.
 */

/* ── 설정 ── */
// 댓글을 저장할 스프레드시트. 비워 두면 이 스크립트가 붙어 있는 시트(「뉴비키 등록 요청」)를 써요.
// 스크립트가 시트에 붙어 있지 않다면 「뉴비키 등록 요청」 주소의 /d/와 /edit 사이 ID를 넣어 주세요.
var CMT_BOOK_ID = "";
// 용어 ID 확인용: 「뉴비키 용어집」 스프레드시트 ID (비워 두면 id 형식만 검사해요)
// 넣으면 그 시트 1행에서 머리글이 "id"인 열을 읽어서, 실제로 있는 용어에만 댓글을 받아요.
var CMT_TERM_BOOK_ID = "";

var CMT_SHEET = "댓글";
var LIKE_SHEET = "좋아요";
var CMT_HEAD = ["댓글ID", "용어ID", "익명번호", "내용", "작성시각", "숨김"];
var LIKE_HEAD = ["댓글ID", "브라우저ID", "시각"];
var CMT_MAX = 500;

/* ── 진입점 ── */
function handleCommentsGet_(e) {
  var p = (e && e.parameter) || {};
  if (p.only !== "comments") return null;
  try {
    var term = String(p.term || "");
    if (!cmtValidTermFormat_(term)) return cmtJson_({ ok: false, error: "term" });
    var browser = cmtValidBrowser_(p.browser) ? String(p.browser) : "";
    return cmtJson_({ ok: true, comments: cmtList_(term, browser) });
  } catch (err) {
    return cmtJson_({ ok: false, error: "server" });
  }
}

function handleCommentPost_(e) {
  var body;
  try { body = JSON.parse((e && e.postData && e.postData.contents) || "{}"); } catch (err) { return null; }
  if (!body || (body.type !== "comment" && body.type !== "like")) return null;
  try {
    return cmtJson_(body.type === "comment" ? cmtAdd_(body) : cmtLike_(body));
  } catch (err) {
    return cmtJson_({ ok: false, error: "server" });
  }
}

/* ── 조회 ── */
function cmtList_(term, browser) {
  var rows = cmtRows_(cmtSheet_(CMT_SHEET, CMT_HEAD));
  var likes = cmtLikeIndex_(browser);
  var out = [];
  rows.forEach(function (r) {
    if (String(r[1]) !== term || cmtIsHidden_(r[5])) return;
    var id = String(r[0]);
    out.push({
      id: id,
      number: Number(r[2]),
      content: String(r[3]),
      createdAt: cmtIso_(r[4]),
      likes: likes.count[id] || 0,
      liked: !!likes.mine[id]
    });
  });
  out.sort(function (a, b) { return a.createdAt < b.createdAt ? 1 : a.createdAt > b.createdAt ? -1 : b.number - a.number; });
  return out;
}

/* ── 작성 ── */
function cmtAdd_(body) {
  var term = String(body.term || "");
  if (!cmtValidTermFormat_(term) || !cmtTermExists_(term)) return { ok: false, error: "term" };
  var content = String(body.content == null ? "" : body.content).trim();
  var len = Array.from(content).length;
  if (len < 1 || len > CMT_MAX) return { ok: false, error: "length" };

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { ok: false, error: "busy" };
  try {
    var sheet = cmtSheet_(CMT_SHEET, CMT_HEAD);
    var max = 0;
    cmtRows_(sheet).forEach(function (r) {
      if (String(r[1]) === term) max = Math.max(max, Number(r[2]) || 0);
    });
    var number = max + 1;
    var id = Utilities.getUuid();
    var now = new Date();
    // =, +, -, @로 시작하면 시트가 수식으로 읽지 않게 앞에 ' 를 붙여 글자로 저장
    var safe = /^[=+\-@]/.test(content) ? "'" + content : content;
    var row = sheet.getLastRow() + 1;
    sheet.getRange(row, 1, 1, 6).setValues([[id, term, number, safe, now, false]]);
    sheet.getRange(row, 6).insertCheckboxes();
    SpreadsheetApp.flush();
    return { ok: true, comment: { id: id, number: number, content: content, createdAt: now.toISOString(), likes: 0, liked: false } };
  } finally {
    lock.releaseLock();
  }
}

/* ── 좋아요 ── */
function cmtLike_(body) {
  var id = String(body.commentId || "");
  var browser = String(body.browser || "");
  if (!id || id.length > 64 || !cmtValidBrowser_(browser)) return { ok: false, error: "bad" };

  var exists = cmtRows_(cmtSheet_(CMT_SHEET, CMT_HEAD)).some(function (r) { return String(r[0]) === id && !cmtIsHidden_(r[5]); });
  if (!exists) return { ok: false, error: "comment" };

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return { ok: false, error: "busy" };
  try {
    var sheet = cmtSheet_(LIKE_SHEET, LIKE_HEAD);
    var rows = cmtRows_(sheet);
    var count = 0, mine = false;
    rows.forEach(function (r) {
      if (String(r[0]) !== id) return;
      count++;
      if (String(r[1]) === browser) mine = true;
    });
    if (!mine) {
      sheet.appendRow([id, browser, new Date()]);
      SpreadsheetApp.flush();
      count++;
    }
    return { ok: true, likes: count, liked: true };
  } finally {
    lock.releaseLock();
  }
}

/* ── 도우미 ── */
function cmtBook_() {
  return CMT_BOOK_ID ? SpreadsheetApp.openById(CMT_BOOK_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

// 탭이 없으면 머리글과 함께 만든다
function cmtSheet_(name, head) {
  var book = cmtBook_();
  var sheet = book.getSheetByName(name);
  if (!sheet) {
    sheet = book.insertSheet(name);
    sheet.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function cmtRows_(sheet) {
  var last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
}

// 좋아요 탭을 한 번 읽어서 댓글별 수와 "이 브라우저가 누른 댓글"을 만든다
function cmtLikeIndex_(browser) {
  var count = {}, mine = {};
  cmtRows_(cmtSheet_(LIKE_SHEET, LIKE_HEAD)).forEach(function (r) {
    var id = String(r[0]);
    count[id] = (count[id] || 0) + 1;
    if (browser && String(r[1]) === browser) mine[id] = true;
  });
  return { count: count, mine: mine };
}

function cmtIsHidden_(v) { return v === true || String(v).toUpperCase() === "TRUE"; }
function cmtIso_(v) { return v instanceof Date ? v.toISOString() : String(v || ""); }
function cmtValidTermFormat_(t) { return /^[a-z0-9]{1,40}$/.test(t); }
function cmtValidBrowser_(b) { return typeof b === "string" && /^[A-Za-z0-9-]{8,36}$/.test(b); }

// 용어집 시트에 있는 id인지 확인 (10분 캐시). CMT_TERM_BOOK_ID가 비어 있으면 형식 검사만.
function cmtTermExists_(term) {
  if (!CMT_TERM_BOOK_ID) return true;
  var cache = CacheService.getScriptCache();
  var cached = cache.get("cmt.termIds");
  var ids;
  if (cached) ids = JSON.parse(cached);
  else {
    ids = [];
    SpreadsheetApp.openById(CMT_TERM_BOOK_ID).getSheets().forEach(function (sh) {
      var last = sh.getLastRow(), cols = sh.getLastColumn();
      if (last < 2 || cols < 1) return;
      var head = sh.getRange(1, 1, 1, cols).getValues()[0].map(function (h) { return String(h).trim().toLowerCase(); });
      var c = head.indexOf("id");
      if (c < 0) return;
      sh.getRange(2, c + 1, last - 1, 1).getValues().forEach(function (r) { if (r[0]) ids.push(String(r[0]).trim()); });
    });
    try { cache.put("cmt.termIds", JSON.stringify(ids), 600); } catch (err) {}
  }
  return ids.indexOf(term) >= 0;
}

function cmtJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
