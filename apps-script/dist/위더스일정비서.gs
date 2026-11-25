/**
 * 위더스 일정비서 - 한 파일 버전 (자동 생성 파일, 직접 고치지 말고 apps-script/ 의 원본을 고친 뒤 build.js 실행)
 * 사용법: 구글 시트 > 확장 프로그램 > Apps Script > Code.gs 내용을 모두 지우고 이 파일 전체를 붙여넣기
 */
/**
 * 위더스 일정비서 - 구글 Apps Script 서버 (무료)
 *
 * 하는 일
 *  1) 안드로이드 문자 전달 앱이 보낸 문자를 받음 (doPost)
 *  2) 문자에서 날짜/시간/작업을 읽고 구글 시트의 일정표와 비교
 *  3) 카카오톡 "나에게 보내기"로 아빠에게 알림 (선택)
 *  4) 휴대폰 화면(일정 달력)을 보여줌 (doGet)
 *
 * 설치 방법은 SETUP.md 를 보세요.
 */

// ===== 작업 기준 (아빠 기준에 맞게 고치세요) =====
var CONFIG = {
  open: 8 * 60,            // 작업 시작 가능 시각 08:00
  close: 19 * 60,          // 작업 끝나야 하는 시각 19:00
  buffer: 30,              // 작업 사이 이동시간(분)
  fullStart: 8 * 60 + 30,  // 입주/이사청소 시작 08:30
  fullDur: 9 * 60,         // 입주/이사청소는 하루 종일(9시간)
  unit: { "벽걸이": 60, "스탠드": 90, "투인원": 150, "시스템": 120, "세탁기": 120 }, // 1대당 분
  closedDays: [0]          // 휴무 요일 (0=일, 6=토)
};
var TZ = "Asia/Seoul";
var SHEET_JOBS = "일정";
var SHEET_INBOX = "문의";
var JOB_COLS = ["id", "date", "start", "dur", "fullDay", "label", "customer", "phone", "address", "inboxId", "createdAt"];
var INBOX_COLS = ["id", "receivedAt", "from", "text", "status", "source", "parsed", "notified"];

var _logic = null;
function logic_() { return _logic || (_logic = makeLogic(CONFIG)); }
function props_() { return PropertiesService.getScriptProperties(); }

/** 서울 시각 기준 '지금' (Date의 연/월/일/시/분이 서울 시각과 같도록 만든 값) */
function seoulNow_() {
  var p = Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd-HH-mm").split("-").map(Number);
  return new Date(p[0], p[1] - 1, p[2], p[3], p[4]);
}
function stamp_() { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd HH:mm"); }

// ======================================================================
// 처음 한 번 실행: setup
// ======================================================================
function setup() {
  var ss = SpreadsheetApp.getActive();
  ensureSheet_(ss, SHEET_JOBS, JOB_COLS);
  ensureSheet_(ss, SHEET_INBOX, INBOX_COLS);
  var p = props_();
  if (!p.getProperty("SECRET")) p.setProperty("SECRET", Utilities.getUuid().replace(/-/g, "").slice(0, 20));
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === "keepKakaoAlive") ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger("keepKakaoAlive").timeBased().everyDays(1).atHour(4).inTimezone(TZ).create();
  var deployed = false;
  try { deployed = !!appUrl_(); } catch (err) {}
  if (deployed) showLinks();
  else Logger.log("준비 완료. 이제 [배포 > 새 배포 > 웹 앱]으로 배포한 뒤 showLinks 를 실행하세요.");
}

function ensureSheet_(ss, name, cols) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.getRange(1, 1, sh.getMaxRows(), cols.length).setNumberFormat("@"); // 날짜/전화번호가 자동 변환되지 않게
  sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight("bold");
  sh.setFrozenRows(1);
  return sh;
}

/** 배포 후 실행: 아빠 휴대폰에 넣을 주소들을 로그에 보여줌 */
function showLinks() {
  var url = appUrl_(), s = secret_();
  if (!url) { Logger.log("아직 배포 전이에요. [배포 > 새 배포 > 웹 앱]을 먼저 하세요."); return; }
  if (!/\/exec$/.test(url)) {
    // 편집기에서 실행하면 구글이 로그인 전용 /dev 주소를 돌려줄 때가 있어요. 그 주소는 아빠 폰에서 안 열려서 적지 않아요.
    var msg = "[배포 > 배포 관리]의 '웹 앱 URL'(…/exec) 뒤에 ?key=" + s + " 를 붙여서 쓰세요";
    Logger.log("편집기에서는 진짜 주소(/exec)를 알 수 없어요. " + msg);
    url = null;
  }
  Logger.log("① 휴대폰 화면 주소 (아빠 폰 크롬에서 열고 '홈 화면에 추가'):\n" + url + "?key=" + s);
  Logger.log("② 문자 전달 앱에 넣을 주소 (Webhook URL):\n" + url + "?key=" + s);
  Logger.log("③ 카카오 Redirect URI (카카오 개발자 사이트에 등록):\n" + url);
  var ss = SpreadsheetApp.getActive(), sh = ss.getSheetByName("안내") || ss.insertSheet("안내", 0);
  sh.clear();
  sh.getRange(1, 1, 6, 2).setValues([
    ["위더스 일정비서 주소", "이 칸의 주소를 복사해서 쓰세요 (key 는 비밀번호와 같아요)"],
    ["① 아빠 폰 화면 주소", url ? url + "?key=" + s : msg],
    ["② 문자 전달 앱 주소", url ? url + "?key=" + s : "① 과 같은 주소 (아빠 폰 화면 맨 아래 '전달 앱 주소 복사' 버튼으로도 복사돼요)"],
    ["③ 카카오 Redirect URI", url || "[배포 관리]의 '웹 앱 URL' 그대로 (…/exec)"],
    ["", ""],
    ["만든 시각", stamp_()]
  ]);
  sh.getRange(1, 1, 1, 2).setFontWeight("bold");
  sh.setColumnWidth(1, 170); sh.setColumnWidth(2, 640);
  Logger.log("시트의 '안내' 탭에도 주소를 적어 두었어요.");
}

function appUrl_() { return props_().getProperty("APP_URL") || ScriptApp.getService().getUrl(); }
function secret_() {
  var s = props_().getProperty("SECRET");
  if (!s) throw new Error("setup 을 먼저 실행하세요.");
  return s;
}
function checkKey_(k) { if (!k || k !== props_().getProperty("SECRET")) throw new Error("주소의 key 가 맞지 않아요."); }

// ======================================================================
// 웹 주소로 들어오는 요청
// ======================================================================
function doGet(e) {
  var q = (e && e.parameter) || {};
  if (q.code && q.state) return kakaoCallback_(q);
  if (!q.key || q.key !== props_().getProperty("SECRET")) return simplePage_("주소가 올바르지 않아요", "처음 받은 주소(?key=... 가 붙은 주소)로 열어 주세요.");
  if (q.kakao === "login") return kakaoLoginPage_();

  var t = typeof INDEX_HTML === "string" ? HtmlService.createTemplate(INDEX_HTML) : HtmlService.createTemplateFromFile("Index");
  t.key = q.key;
  t.appUrl = appUrl_();
  t.cfgJson = JSON.stringify(CONFIG);
  t.logicSrc = makeLogic.toString();
  t.bootJson = JSON.stringify(state_()).replace(/</g, "\\u003c");
  return t.evaluate()
    .setTitle("위더스 일정비서")
    .addMetaTag("viewport", "width=device-width, initial-scale=1, viewport-fit=cover");
}

/** 문자 전달 앱이 보내는 요청 */
function doPost(e) {
  var q = (e && e.parameter) || {};
  if (!q.key || q.key !== props_().getProperty("SECRET")) return json_({ ok: false, error: "bad key" });
  props_().setProperty("LAST_POST_AT", stamp_()); // 문자 전달 앱 연결 확인용
  var body = {};
  try { body = JSON.parse((e.postData && e.postData.contents) || "{}"); } catch (err) { body = q; }
  var from = String(body.from || body.sender || body.phone || body.number || q.from || "").trim();
  var text = String(body.text || body.message || body.body || body.msg || q.text || "").trim();
  if (!text) return json_({ ok: true, skipped: "empty" });
  var res = ingest_(from, text, "sms");
  return json_({ ok: true, result: res.result });
}

function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function simplePage_(title, msg) {
  return HtmlService.createHtmlOutput('<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<div style="font-family:sans-serif;padding:24px;line-height:1.6"><h2>' + title + "</h2><p>" + msg + "</p></div>").setTitle("위더스 일정비서");
}

// ======================================================================
// 문자 처리
// ======================================================================
/**
 * 문자 한 통을 처리해요.
 * source: "sms"(자동 전달) / "manual"(화면에 직접 붙여넣기)
 */
function ingest_(from, text, source) {
  var L = logic_(), now = seoulNow_();
  var id = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, from + "|" + text, Utilities.Charset.UTF_8)
    .map(function (b) { return ((b + 256) % 256).toString(16).padStart(2, "0"); }).join("").slice(0, 16);

  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var existing = readRows_(SHEET_INBOX, INBOX_COLS).filter(function (r) { return r.id === id; })[0];
    if (existing) return { result: "duplicate", id: id }; // 전달 앱이 같은 문자를 다시 보낸 경우
  } finally { lock.releaseLock(); }

  if (source === "sms" && !L.prefilter(text)) return { result: "ignored", id: id };

  var parsed = null;
  if (aiEnabled_()) parsed = aiParse_(text, now);   // AI 켜져 있으면 AI로 (실패하면 규칙으로)
  if (!parsed) parsed = L.parse(text, now);
  if (!parsed.phone && /^\+?\d[\d-]{7,}$/.test(from)) parsed.phone = prettyPhone_(from);
  if (source === "sms" && !parsed.isJob) return { result: "ignored", id: id };

  lock.waitLock(20000);
  try {
    appendRow_(SHEET_INBOX, INBOX_COLS, {
      id: id, receivedAt: stamp_(), from: /^\+?\d[\d-]{7,}$/.test(from) ? prettyPhone_(from) : (from || "직접 입력"), text: text, status: "new",
      source: source, parsed: JSON.stringify(parsed), notified: ""
    });
  } finally { lock.releaseLock(); }

  if (source === "sms" && kakaoConnected_()) {
    try {
      var ev = L.evaluate(parsed, readJobs_(), now);
      if (kakaoSend_(kakaoText_(parsed, ev, text), appUrl_() + "?key=" + secret_())) updateInbox_(id, { notified: "Y" });
    } catch (err) { console.error("카카오 알림 실패: " + err); }
  }
  return { result: "saved", id: id };
}

function prettyPhone_(p) {
  var d = String(p).replace(/[^\d+]/g, "").replace(/^\+82/, "0");
  var m = d.match(/^(01\d)(\d{3,4})(\d{4})$/);
  return m ? m[1] + "-" + m[2] + "-" + m[3] : d;
}

function kakaoText_(r, ev, text) {
  var L = logic_(), head;
  if (ev.status === "conflict") head = "[일정 겹침] 새 문의";
  else if (ev.status === "ok") head = "[가능] 새 문의";
  else if (ev.status === "suggest") head = "[시간 추천] 새 문의";
  else head = "[확인 필요] 새 문의";
  var lines = [head];
  if (r.date) lines.push(L.fmtShort(r.date) + " " + (r.full ? "종일" : r.start != null ? L.fmtTime(r.start) : (r.timeText || "시간 미정")) + " · " + L.shortLabel(r.label));
  else lines.push("날짜 없음 · " + L.shortLabel(r.label));
  lines.push(((r.customer || "") + " " + (r.phone || "")).trim() || "고객 정보 없음");
  if (ev.status === "conflict" && ev.conflicts.length) {
    lines.push("겹침: " + ev.conflicts.slice(0, 2).map(function (j) { return (j.fullDay ? "종일" : L.hm(j.start)) + " " + (j.customer || L.shortLabel(j.label)); }).join(", "));
  }
  if (ev.status === "suggest") lines.push(L.fmtTime(ev.start) + " 비어 있어요");
  if (ev.status === "ok") lines.push("그 시간 비어 있어요");
  if (ev.alts.length) lines.push("가능: " + ev.alts.slice(0, 2).map(function (a) { return L.fmtShort(a.date) + (r.full ? "" : " " + L.fmtTime(a.start)); }).join(" / "));
  if (ev.status === "nodate" || ev.status === "past") lines.push(text.slice(0, 50));
  var s = lines.join("\n");
  return s.length > 200 ? s.slice(0, 199) + "…" : s; // 카카오 텍스트는 200자까지
}

// ======================================================================
// 화면(Index.html)에서 부르는 함수들 - 모두 key 확인
// ======================================================================
function getState(k) { checkKey_(k); return state_(); }

function ingestManual(k, text) {
  checkKey_(k);
  text = String(text || "").trim();
  if (!text) throw new Error("내용이 비어 있어요.");
  var res = ingest_("", text, "manual");
  if (res.result === "duplicate") updateInbox_(res.id, { status: "new" });
  return { id: res.id, state: state_() };
}

function addJob(k, job, inboxId) {
  checkKey_(k);
  if (!job || !/^\d{4}-\d{2}-\d{2}$/.test(job.date)) throw new Error("날짜가 올바르지 않아요.");
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    appendRow_(SHEET_JOBS, JOB_COLS, {
      id: "j" + Date.now(), date: job.date, start: String(job.start | 0), dur: String(job.dur | 0), fullDay: job.fullDay ? "Y" : "",
      label: job.label || "", customer: job.customer || "", phone: job.phone || "", address: job.address || "",
      inboxId: inboxId || "", createdAt: stamp_()
    });
    if (inboxId) updateInbox_(inboxId, { status: "done" });
  } finally { lock.releaseLock(); }
  return state_();
}

function deleteJob(k, id) {
  checkKey_(k);
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_JOBS), vals = sh.getDataRange().getDisplayValues();
    for (var i = vals.length - 1; i >= 1; i--) if (vals[i][0] === id) sh.deleteRow(i + 1);
  } finally { lock.releaseLock(); }
  return state_();
}

function setInboxStatus(k, id, status) {
  checkKey_(k);
  if (["new", "done", "dismissed"].indexOf(status) < 0) throw new Error("bad status");
  updateInbox_(id, { status: status });
  return state_();
}

function sendTestKakao(k) {
  checkKey_(k);
  if (!kakaoConnected_()) throw new Error("카카오가 아직 연결되지 않았어요.");
  return kakaoSend_("[테스트] 위더스 일정비서 알림이 잘 와요.", appUrl_() + "?key=" + secret_());
}

function state_() {
  var L = logic_(), now = seoulNow_(), from = L.key(L.addDays(now, -35));
  var inbox = readRows_(SHEET_INBOX, INBOX_COLS)
    .filter(function (r) { return r.status !== "dismissed"; })
    .reverse().slice(0, 60)
    .map(function (r) { var p = null; try { p = JSON.parse(r.parsed); } catch (e) {} return { id: r.id, receivedAt: r.receivedAt, from: r.from, text: r.text, status: r.status, source: r.source, parsed: p }; });
  return {
    jobs: readJobs_().filter(function (j) { return j.date >= from; }),
    inbox: inbox,
    settings: { kakaoReady: !!props_().getProperty("KAKAO_REST_KEY"), kakao: kakaoConnected_(), ai: aiEnabled_(), aiModel: aiModel_(), lastPostAt: props_().getProperty("LAST_POST_AT") || "" }
  };
}

// ======================================================================
// 시트 읽기/쓰기
// ======================================================================
function readRows_(name, cols) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error("'" + name + "' 시트가 없어요. setup 을 먼저 실행하세요.");
  var vals = sh.getDataRange().getDisplayValues();
  return vals.slice(1).filter(function (r) { return r[0]; }).map(function (r) {
    var o = {}; cols.forEach(function (c, i) { o[c] = r[i] || ""; }); return o;
  });
}
function readJobs_() {
  return readRows_(SHEET_JOBS, JOB_COLS).map(function (r) {
    return { id: r.id, date: r.date, start: Number(r.start) || 0, dur: Number(r.dur) || 60, fullDay: r.fullDay === "Y",
      label: r.label, customer: r.customer, phone: r.phone, address: r.address, inboxId: r.inboxId };
  });
}
function appendRow_(name, cols, o) {
  var sh = SpreadsheetApp.getActive().getSheetByName(name);
  var row = cols.map(function (c) { return o[c] == null ? "" : String(o[c]); });
  var r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, cols.length).setNumberFormat("@").setValues([row]);
}
function updateInbox_(id, patch) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SHEET_INBOX), vals = sh.getDataRange().getDisplayValues();
  for (var i = 1; i < vals.length; i++) {
    if (vals[i][0] !== id) continue;
    Object.keys(patch).forEach(function (c) { var col = INBOX_COLS.indexOf(c); if (col >= 0) sh.getRange(i + 1, col + 1).setValue(String(patch[c])); });
  }
}

// ======================================================================
// 카카오톡 "나에게 보내기" (무료)
// ======================================================================
function kakaoConnected_() { return !!props_().getProperty("KAKAO_REFRESH_TOKEN"); }

function kakaoLoginPage_() {
  var p = props_(), rest = p.getProperty("KAKAO_REST_KEY");
  if (!rest) return simplePage_("카카오 키가 없어요", "Apps Script [프로젝트 설정 > 스크립트 속성]에 KAKAO_REST_KEY 를 먼저 넣어 주세요.");
  var url = "https://kauth.kakao.com/oauth/authorize?response_type=code&scope=talk_message" +
    "&client_id=" + encodeURIComponent(rest) + "&redirect_uri=" + encodeURIComponent(appUrl_()) + "&state=" + encodeURIComponent(secret_());
  return HtmlService.createHtmlOutput('<meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<div style="font-family:sans-serif;padding:24px;line-height:1.6"><h2>카카오톡 알림 연결</h2>' +
    "<p>아빠 카카오 계정으로 로그인하고 '카카오톡 메시지 전송'에 동의하면 끝이에요.</p>" +
    '<p><a href="' + url + '" target="_top" style="display:inline-block;background:#fee500;color:#191919;padding:12px 18px;border-radius:10px;font-weight:700;text-decoration:none">카카오로 로그인</a></p></div>')
    .setTitle("카카오 연결");
}

function kakaoCallback_(q) {
  if (q.state !== props_().getProperty("SECRET")) return simplePage_("연결 실패", "주소가 올바르지 않아요.");
  try {
    kakaoTokenRequest_({ grant_type: "authorization_code", code: q.code, redirect_uri: appUrl_() });
    kakaoSend_("[연결 완료] 이제 새 문의가 오면 여기로 알려드릴게요.", appUrl_() + "?key=" + secret_());
    return simplePage_("카카오톡 연결 완료", "카카오톡 '나와의 채팅'에 테스트 메시지를 보냈어요. 이 창은 닫아도 돼요.");
  } catch (err) {
    return simplePage_("연결 실패", String(err));
  }
}

function kakaoTokenRequest_(params) {
  var p = props_();
  params.client_id = p.getProperty("KAKAO_REST_KEY");
  var cs = p.getProperty("KAKAO_CLIENT_SECRET");
  if (cs) params.client_secret = cs;
  var res = UrlFetchApp.fetch("https://kauth.kakao.com/oauth/token", { method: "post", payload: params, muteHttpExceptions: true });
  var data = JSON.parse(res.getContentText() || "{}");
  if (res.getResponseCode() !== 200 || !data.access_token) throw new Error("카카오 토큰 오류: " + res.getContentText());
  p.setProperty("KAKAO_ACCESS_TOKEN", data.access_token);
  p.setProperty("KAKAO_ACCESS_EXP", String(Date.now() + (data.expires_in || 3600) * 1000));
  if (data.refresh_token) p.setProperty("KAKAO_REFRESH_TOKEN", data.refresh_token); // 만료가 가까우면 새로 줌
  return data.access_token;
}

function kakaoAccessToken_(force) {
  var p = props_(), tok = p.getProperty("KAKAO_ACCESS_TOKEN"), exp = Number(p.getProperty("KAKAO_ACCESS_EXP") || 0);
  if (!force && tok && exp > Date.now() + 60000) return tok;
  return kakaoTokenRequest_({ grant_type: "refresh_token", refresh_token: p.getProperty("KAKAO_REFRESH_TOKEN") });
}

function kakaoSend_(text, link) {
  var tpl = { object_type: "text", text: text, link: { web_url: link, mobile_web_url: link }, button_title: "일정 확인" };
  for (var attempt = 0; attempt < 2; attempt++) {
    var res = UrlFetchApp.fetch("https://kapi.kakao.com/v2/api/talk/memo/default/send", {
      method: "post", headers: { Authorization: "Bearer " + kakaoAccessToken_(attempt > 0) },
      payload: { template_object: JSON.stringify(tpl) }, muteHttpExceptions: true
    });
    if (res.getResponseCode() === 200) return true;
    if (res.getResponseCode() !== 401) { console.error("카카오 전송 실패: " + res.getContentText()); return false; }
  }
  return false;
}

/** 매일 새벽 자동 실행: 카카오 연결이 끊기지 않게 토큰을 갱신 */
function keepKakaoAlive() {
  if (kakaoConnected_()) kakaoAccessToken_(true);
}


/**
 * 일정 판단 로직 (서버와 휴대폰 화면이 같은 코드를 씀)
 *
 * makeLogic 안에는 바깥 변수를 쓰면 안 돼요.
 * 이 함수의 소스코드가 그대로 화면(Index.html)에도 들어가기 때문이에요.
 * 날짜는 모두 'YYYY-MM-DD' 문자열, 시간은 자정부터 센 '분'(예: 09:30 → 570)으로 다뤄요.
 */
function makeLogic(cfg) {
  "use strict";
  var OPEN = cfg.open, CLOSE = cfg.close, BUF = cfg.buffer;
  var FULL_START = cfg.fullStart, FULL_DUR = cfg.fullDur;
  var UNIT = cfg.unit, CLOSED = cfg.closedDays || [0];
  var DOW = "일월화수목금토";

  // ---------- 날짜/시간 도우미 ----------
  function pad(n) { return String(n).padStart(2, "0"); }
  function key(d) { return d.getFullYear() + "-" + pad(d.getMonth() + 1) + "-" + pad(d.getDate()); }
  function fromKey(k) { var p = k.split("-").map(Number); return new Date(p[0], p[1] - 1, p[2]); }
  function addDays(d, n) { var x = new Date(d.getTime()); x.setDate(x.getDate() + n); return x; }
  function day0(now) { return new Date(now.getFullYear(), now.getMonth(), now.getDate()); }
  function minsOf(now) { return now.getHours() * 60 + now.getMinutes(); }
  function mondayOf(d) { return addDays(d, -((d.getDay() + 6) % 7)); }
  function isClosed(k) { return CLOSED.indexOf(fromKey(k).getDay()) >= 0; }
  function fmtDate(k) { var d = fromKey(k); return (d.getMonth() + 1) + "월 " + d.getDate() + "일(" + DOW[d.getDay()] + ")"; }
  function fmtShort(k) { var d = fromKey(k); return (d.getMonth() + 1) + "/" + d.getDate() + "(" + DOW[d.getDay()] + ")"; }
  function fmtTime(min) { var h = Math.floor(min / 60), m = min % 60; return (h < 12 ? "오전 " : "오후 ") + (h % 12 || 12) + "시" + (m ? " " + m + "분" : ""); }
  function hm(min) { return pad(Math.floor(min / 60)) + ":" + pad(min % 60); }
  function fmtDur(d) { var h = Math.floor(d / 60), m = d % 60; return (h ? h + "시간" : "") + (m ? (h ? " " : "") + m + "분" : ""); }
  function whenText(date, start, full) { return fmtDate(date) + " " + (full ? "하루 종일" : fmtTime(start)); }

  // ---------- 작업 이름/소요시간 ----------
  function labelFor(kind, items, pyeong) {
    var py = pyeong ? " · " + pyeong + "평" : "";
    if (kind === "move_in") return "입주청소" + py;
    if (kind === "move_out") return "이사청소" + py;
    if (kind === "home") return "거주청소" + py;
    if (kind === "washer") return "세탁기 청소";
    if (kind === "aircon" && items.length) return "에어컨 청소 · " + items.join(", ");
    return "청소 (종류 확인 필요)";
  }
  function shortLabel(l) {
    var i = l.indexOf("·");
    var core = /^에어컨/.test(l) && i > 0 ? l.slice(i + 1).trim() : (i > 0 ? l.slice(0, i).trim() : l);
    return core.replace(/\s*(\d+)대/g, "$1").replace(/, /g, "·");
  }

  // ---------- 업무 문자인지 대충 거르기 ----------
  var NEGATIVE = /인증번호|\[Web발신\]|\(광고\)|무료수신거부|출금|입금|승인|결제|잔액|택배|배송|080-\d/;
  var KEYWORD = /청소|에어컨|에어콘|세척|입주|이사|견적|예약|방문|벽걸이|스탠드|시스템/;
  function prefilter(text) {
    if (NEGATIVE.test(text)) return false;
    return KEYWORD.test(text) || /\d|오늘|내일|모레|요일/.test(text);
  }

  // ---------- 규칙으로 문자 읽기 (무료) ----------
  function parse(raw, now) {
    var T = day0(now), t = String(raw).replace(/\s+/g, " "), m;
    var r = { via: "rule", isJob: false, date: null, dateText: "", start: null, timeText: "", win: null, full: false,
      dur: 60, label: "", customer: "", phone: "", address: "" };

    m = t.match(/01[016789][-\s.]?\d{3,4}[-\s.]?\d{4}/);
    if (m) { r.phone = m[0].replace(/[\s.]/g, "-"); t = t.replace(m[0], " "); }

    var rel = { "오늘": 0, "금일": 0, "내일": 1, "낼": 1, "모레": 2, "글피": 3 };
    function setMD(mo, d, txt) {
      var dt = new Date(T.getFullYear(), mo - 1, d);
      if (dt < addDays(T, -30)) dt = new Date(T.getFullYear() + 1, mo - 1, d);
      if (dt.getMonth() === mo - 1) { r.date = key(dt); r.dateText = txt; }
    }
    if ((m = t.match(/(오늘|금일|내일|모레|글피)/))) { r.date = key(addDays(T, rel[m[1]])); r.dateText = m[1]; }
    else if ((m = t.match(/(이번\s*주|금주|다다음\s*주|다음\s*주|담주)?\s*([월화수목금토일])\s*요일/))) {
      var idx = DOW.indexOf(m[2]), off = (idx + 6) % 7, w = (m[1] || "").replace(/\s/g, ""), mon = mondayOf(T), d;
      if (w === "다음주" || w === "담주") d = addDays(mon, 7 + off);
      else if (w === "다다음주") d = addDays(mon, 14 + off);
      else if (w === "이번주" || w === "금주") d = addDays(mon, off);
      else d = addDays(T, ((idx - T.getDay() + 7) % 7) || 7);
      r.date = key(d); r.dateText = m[0].trim();
    }
    else if ((m = t.match(/(\d{1,2})\s*월\s*(\d{1,2})\s*일/))) setMD(+m[1], +m[2], m[0]);
    else if ((m = t.match(/(?:^|[^\d])(\d{1,2})\s*[\/.]\s*(\d{1,2})(?!\d)/)) && +m[1] >= 1 && +m[1] <= 12 && +m[2] >= 1 && +m[2] <= 31) setMD(+m[1], +m[2], m[1] + "/" + m[2]);
    else if ((m = t.match(/(?:^|[^\d월])(\d{1,2})\s*일(?![\d요])/))) {
      var dd = +m[1], dt = new Date(T.getFullYear(), T.getMonth(), dd);
      if (dt < T) dt = new Date(T.getFullYear(), T.getMonth() + 1, dd);
      if (dd >= 1 && dd <= 31) { r.date = key(dt); r.dateText = m[1] + "일"; }
    }

    if ((m = t.match(/(오전|오후|아침|낮|저녁|밤)?\s*(\d{1,2})\s*시(?!스템|간)\s*(반|(\d{1,2})\s*분)?/))) {
      var h = +m[2], mi = m[3] === "반" ? 30 : (m[4] ? +m[4] : 0), ap = m[1] || "";
      if (/오후|저녁|밤/.test(ap) && h < 12) h += 12;
      else if (ap === "낮" && h <= 6) h += 12;
      else if (!ap && h >= 1 && h <= 7) h += 12;
      r.start = h * 60 + mi; r.timeText = m[0].trim();
    } else if ((m = t.match(/(?:^|[^\d])(\d{1,2}):(\d{2})/))) {
      r.start = +m[1] * 60 + +m[2]; r.timeText = m[1] + ":" + m[2];
    } else if (/오전/.test(t)) { r.win = { from: 540, to: 690 }; r.timeText = "오전 중"; }
    else if (/오후|점심\s*이후/.test(t)) { r.win = { from: 780, to: 1050 }; r.timeText = "오후 중"; }
    else if (/저녁|퇴근/.test(t)) { r.win = { from: 1020, to: 1080 }; r.timeText = "저녁"; }
    else { r.win = { from: 540, to: 1080 }; r.timeText = ""; }

    var items = [], dur = 0, kind = "other";
    [["시스템", /시스템|천장형|천정형/], ["투인원", /투인원|2\s?in\s?1/i], ["스탠드", /스탠드/], ["벽걸이", /벽걸이/]].forEach(function (k) {
      var km = t.match(k[1]); if (!km) return;
      var after = t.slice(km.index + km[0].length), cm = after.match(/^[^\d.,!?]{0,10}?(\d+)\s*대/);
      var n = cm ? +cm[1] : 1; items.push(k[0] + " " + n + "대"); dur += UNIT[k[0]] * n;
    });
    if (!items.length && /에어컨|에어콘/.test(t)) { var cm2 = t.match(/(\d+)\s*대/), n2 = cm2 ? +cm2[1] : 1; items.push("벽걸이 " + n2 + "대"); dur += UNIT["벽걸이"] * n2; }
    var py = t.match(/(\d+)\s*평/), pyeong = py ? +py[1] : null;
    if (/(입주|이사)\s*청소/.test(t)) { kind = /입주/.test(t) ? "move_in" : "move_out"; r.full = true; dur = FULL_DUR; }
    else if (/거주\s*청소|집\s*청소|(?<!\d\s*)대청소/.test(t)) { kind = "home"; dur = Math.max(dur, 300); }
    else if (/세탁기/.test(t)) { kind = "washer"; dur = Math.max(dur, UNIT["세탁기"]); }
    else if (items.length) kind = "aircon";
    else dur = 120;
    r.label = labelFor(kind, items, pyeong);
    r.dur = Math.max(60, dur);
    if (r.full) { r.start = FULL_START; r.win = null; r.timeText = r.timeText || "하루 종일"; }

    var nm = t.match(/(?:이름은|성함은|저는|고객명\s*:?)\s*([가-힣]{2,4}?)(?=입니다|이에요|예요|이요|[\s.,!~]|$)/) ||
             t.match(/(?:^|[\s.,~!])([가-힣]{2,4}?)(?=(?:입니다|이에요|예요)(?:[\s.,!~]|$))/);
    if (nm && !/사무실|아파트|가게|매장|빌라|주택|상가|오피스텔|학원|식당/.test(nm[1])) r.customer = nm[1];

    var am = t.match(/(?:[가-힣]+(?:특별시|광역시|시|도)\s)?[가-힣]+(?:구|군)\s[가-힣0-9]+(?:동|읍|면|로|길)(?:\s+[A-Za-z0-9가-힣\-]+){0,5}/);
    if (am) {
      var toks = am[0].split(/\s+/), keep = [], okTok = /(\d|[A-Za-z]|동|호|층|타운|아파트|빌라|자이|힐즈|파크|캐슬|뷰|오피스텔|빌딩|타워|센터|마을|단지|힐스|푸르지오|래미안|시티)$/;
      for (var i = 0; i < toks.length; i++) {
        var tk = toks[i], stripped = tk.replace(/(입니다|이고|이에요|예요|이요|이라서?|에서|이며|고)$/, "");
        if (i < 3) { keep.push(stripped); if (stripped !== tk) break; continue; }
        if (!okTok.test(stripped)) break;
        keep.push(stripped); if (stripped !== tk) break;
      }
      r.address = keep.join(" ");
    }

    r.isJob = !NEGATIVE.test(t) && (KEYWORD.test(t) || (!!r.date && (r.start != null || /가능|되나요|될까요|되세요/.test(t))));
    return r;
  }

  // ---------- AI 결과를 같은 모양으로 바꾸기 ----------
  function fromAI(o, now) {
    var r = { via: "ai", isJob: !!o.is_job_request, date: null, dateText: o.date_text || "", start: null, timeText: "", win: null,
      full: false, dur: 60, label: "", customer: o.customer_name || "", phone: o.phone || "", address: o.address || "" };
    if (o.date && /^\d{4}-\d{2}-\d{2}$/.test(o.date)) r.date = o.date;
    var tm = o.start_time && String(o.start_time).match(/^(\d{1,2}):(\d{2})/);
    if (tm) { r.start = +tm[1] * 60 + +tm[2]; r.timeText = tm[1] + ":" + tm[2]; }
    else {
      var W = { morning: [540, 690, "오전 중"], afternoon: [780, 1050, "오후 중"], evening: [1020, 1080, "저녁"] }[o.time_window];
      r.win = W ? { from: W[0], to: W[1] } : { from: 540, to: 1080 }; r.timeText = W ? W[2] : "";
    }
    var names = { wall: "벽걸이", stand: "스탠드", two_in_one: "투인원", system: "시스템" }, items = [], dur = 0;
    (o.units || []).forEach(function (u) { var n = Math.max(1, u.count | 0), k = names[u.kind]; if (!k) return; items.push(k + " " + n + "대"); dur += UNIT[k] * n; });
    var kind = o.job_type;
    if (kind === "aircon" && !items.length) { items.push("벽걸이 1대"); dur = UNIT["벽걸이"]; }
    if (kind === "move_in" || kind === "move_out") { r.full = true; dur = FULL_DUR; r.start = FULL_START; r.win = null; r.timeText = "하루 종일"; }
    else if (kind === "home") dur = Math.max(dur, 300);
    else if (kind === "washer") dur = Math.max(dur, UNIT["세탁기"]);
    else if (kind !== "aircon") dur = 120;
    r.label = labelFor(kind, items, o.pyeong);
    r.dur = Math.max(60, dur);
    return r;
  }

  // ---------- 일정표와 비교 ----------
  function jobsOn(jobs, k) { return jobs.filter(function (j) { return j.date === k; }).sort(function (a, b) { return a.start - b.start; }); }
  function clash(jobs, k, start, dur, full) {
    var js = jobsOn(jobs, k);
    if (full) return js;
    return js.filter(function (j) { return j.fullDay || (start < j.start + j.dur + BUF && j.start < start + dur + BUF); });
  }
  function fits(jobs, k, start, dur, full, now) {
    if (k === key(day0(now)) && start < minsOf(now) + 60) return false;
    if (!full && (start < OPEN || start + dur > CLOSE)) return false;
    return clash(jobs, k, start, dur, full).length === 0;
  }
  function bestStart(jobs, k, dur, pref, lo, hi, skip, now) {
    var c = [];
    for (var s = lo; s <= hi; s += 30) if (s !== skip) c.push(s);
    c.sort(function (a, b) { return Math.abs(a - pref) - Math.abs(b - pref) || a - b; });
    for (var i = 0; i < c.length; i++) if (fits(jobs, k, c[i], dur, false, now)) return c[i];
    return null;
  }
  function findAlts(r, jobs, n, now) {
    var T = day0(now), tk = key(T), out = [];
    var base = r.date && r.date >= tk ? fromKey(r.date) : addDays(T, 1);
    var pref = r.start != null ? r.start : (r.win ? r.win.from : 600);
    for (var i = 0; i < 40 && out.length < n; i++) {
      var k = key(addDays(base, i));
      if (isClosed(k) || k < tk) continue;
      if (r.full) { if (k !== r.date && fits(jobs, k, FULL_START, FULL_DUR, true, now)) out.push({ date: k, start: FULL_START, sameDay: false }); continue; }
      var s = bestStart(jobs, k, r.dur, pref, OPEN, CLOSE - r.dur, k === r.date ? r.start : null, now);
      if (s != null) out.push({ date: k, start: s, sameDay: k === r.date });
    }
    return out;
  }
  function evaluate(r, jobs, now) {
    var tk = key(day0(now)), ev = { status: "", conflicts: [], start: r.start, alts: [], windowFull: false };
    if (!r.date) ev.status = "nodate";
    else if (r.date < tk || (r.date === tk && r.start != null && !r.full && r.start < minsOf(now))) ev.status = "past";
    else if (isClosed(r.date)) ev.status = "closed";
    else if (r.full) { ev.conflicts = clash(jobs, r.date, 0, 0, true); ev.status = ev.conflicts.length ? "conflict" : "ok"; ev.start = FULL_START; }
    else if (r.start != null) {
      if (r.start < OPEN || r.start + r.dur > CLOSE) ev.status = "outside";
      else { ev.conflicts = clash(jobs, r.date, r.start, r.dur, false); ev.status = ev.conflicts.length ? "conflict" : "ok"; }
    } else {
      var w = r.win || { from: 540, to: 1080 };
      var s = bestStart(jobs, r.date, r.dur, w.from, w.from, Math.min(w.to, CLOSE - r.dur), null, now);
      if (s != null) { ev.status = "suggest"; ev.start = s; }
      else { ev.status = "conflict"; ev.conflicts = jobsOn(jobs, r.date); ev.windowFull = true; }
    }
    if (ev.status !== "ok" && ev.status !== "suggest") ev.alts = findAlts(r, jobs, 3, now);
    return ev;
  }
  function pill(status) {
    if (status === "ok") return { tone: "ok", text: "가능" };
    if (status === "suggest") return { tone: "ok", text: "시간 추천" };
    if (status === "conflict") return { tone: "bad", text: "일정 겹침" };
    return { tone: "warn", text: "확인 필요" };
  }

  // ---------- 고객 답장 ----------
  function replyText(r, ev, sel, added) {
    var head = "안녕하세요 " + (r.customer ? r.customer + "님, " : "") + "위더스환경입니다.\n";
    if (sel) {
      return head + whenText(sel.date, sel.start, r.full) + " " + r.label +
        (added ? " 예약 확정되었습니다.\n방문 전날 다시 연락드리겠습니다. 감사합니다." : " 예약 가능합니다. 이 시간으로 잡아드릴까요?");
    }
    if (ev.status === "nodate") return head + "문의 감사합니다. 원하시는 날짜와 시간을 알려주시면 바로 확인해 드리겠습니다.";
    var why = ev.status === "closed" ? "휴무일이라" : ev.status === "outside" ? "작업 가능 시간(" + fmtTime(OPEN) + "~" + fmtTime(CLOSE) + ")이 아니라" : ev.status === "past" ? "이미 지난 시간이라" : "이미 다른 예약이 있어서";
    var asked = r.date ? fmtDate(r.date) + (r.full ? "" : " " + (r.start != null ? fmtTime(r.start) : r.timeText || "")) : "말씀하신 날짜";
    var t = head + "말씀하신 " + asked.trim() + "에는 " + why + " 어렵습니다.";
    if (ev.alts.length) {
      t += "\n가능한 시간은\n" + ev.alts.map(function (a, i) { return "①②③"[i] + " " + whenText(a.date, a.start, r.full); }).join("\n") +
        "\n입니다. 편하신 시간 알려주시면 바로 잡아드릴게요.";
    }
    return t;
  }

  return {
    OPEN: OPEN, CLOSE: CLOSE, BUF: BUF, FULL_START: FULL_START, FULL_DUR: FULL_DUR, DOW: DOW,
    pad: pad, key: key, fromKey: fromKey, addDays: addDays, day0: day0, minsOf: minsOf, mondayOf: mondayOf, isClosed: isClosed,
    fmtDate: fmtDate, fmtShort: fmtShort, fmtTime: fmtTime, hm: hm, fmtDur: fmtDur, whenText: whenText, shortLabel: shortLabel,
    prefilter: prefilter, parse: parse, fromAI: fromAI, jobsOn: jobsOn, evaluate: evaluate, pill: pill, replyText: replyText
  };
}


/**
 * (선택) AI로 문자 읽기 - Claude API
 *
 * 스크립트 속성에 ANTHROPIC_API_KEY 가 없으면 AI는 꺼져 있고, 무료 규칙 방식으로 읽어요.
 * 키를 넣으면 말투가 달라도 더 잘 알아듣지만, 문자 한 통마다 사용료가 들어요 (SETUP.md 참고).
 * AI 호출이 실패하면 자동으로 규칙 방식으로 돌아가요.
 *
 * 스크립트 속성
 *   ANTHROPIC_API_KEY  : Claude API 키
 *   AI_MODEL           : 사용할 모델 (비워두면 claude-opus-5-5)
 */

var AI_DEFAULT_MODEL = "claude-opus-5-5";

var AI_SCHEMA = {
  type: "object",
  additionalProperties: false,
  required: ["is_job_request", "date", "date_text", "start_time", "time_window", "job_type", "units", "pyeong", "customer_name", "phone", "address"],
  properties: {
    is_job_request: { type: "boolean", description: "청소 예약/문의/일정 변경 요청이면 true. 광고, 은행, 택배, 인증번호, 개인적인 대화면 false." },
    date: { anyOf: [{ type: "string", format: "date" }, { type: "null" }], description: "작업 희망 날짜 YYYY-MM-DD. 없으면 null." },
    date_text: { type: "string", description: "문자에서 날짜를 말한 원래 표현. 없으면 빈 문자열." },
    start_time: { anyOf: [{ type: "string" }, { type: "null" }], description: "정확한 시작 시각 HH:MM (24시간). '오전 중'처럼 범위만 있으면 null." },
    time_window: { type: "string", enum: ["none", "morning", "afternoon", "evening"], description: "정확한 시각 없이 범위만 말한 경우의 범위." },
    job_type: { type: "string", enum: ["aircon", "move_in", "move_out", "home", "washer", "other"], description: "aircon=에어컨 청소, move_in=입주청소, move_out=이사청소, home=거주/집 청소, washer=세탁기 청소" },
    units: {
      type: "array",
      description: "에어컨 종류별 대수. 에어컨이 아니면 빈 배열.",
      items: {
        type: "object", additionalProperties: false, required: ["kind", "count"],
        properties: {
          kind: { type: "string", enum: ["wall", "stand", "two_in_one", "system"], description: "wall=벽걸이, stand=스탠드, two_in_one=투인원(2in1), system=시스템/천장형" },
          count: { type: "integer" }
        }
      }
    },
    pyeong: { anyOf: [{ type: "integer" }, { type: "null" }], description: "집 평수. 없으면 null." },
    customer_name: { anyOf: [{ type: "string" }, { type: "null" }] },
    phone: { anyOf: [{ type: "string" }, { type: "null" }], description: "문자 본문에 적힌 연락처. 없으면 null." },
    address: { anyOf: [{ type: "string" }, { type: "null" }], description: "작업 주소. 없으면 null." }
  }
};

function aiEnabled_() { return !!props_().getProperty("ANTHROPIC_API_KEY"); }
function aiModel_() { return props_().getProperty("AI_MODEL") || AI_DEFAULT_MODEL; }

function aiSystem_(now) {
  var L = logic_(), today = L.key(now);
  return [
    "당신은 에어컨 청소와 집 청소를 하는 1인 사업자(위더스환경)의 휴대폰으로 온 문자를 읽고, 예약 요청 내용을 정리하는 비서입니다.",
    "오늘은 " + today + " (" + L.DOW[now.getDay()] + "요일)이고, 한국 시간 기준입니다.",
    "규칙:",
    "- '내일', '다음주 수요일', '28일' 같은 표현은 오늘 기준으로 실제 날짜로 바꾸세요. '다음주'는 월요일 시작 기준 다음 주입니다.",
    "- 오전/오후 없이 1~7시라고 하면 오후로 봅니다.",
    "- 문자에 없는 정보는 지어내지 말고 null 또는 빈 값으로 두세요.",
    "- 여러 날짜를 말하면 고객이 가장 원하는 첫 번째 날짜를 고르세요."
  ].join("\n");
}

/** 성공하면 Logic의 parse()와 같은 모양을 돌려주고, 실패하면 null */
function aiParse_(text, now) {
  var p = props_(), model = aiModel_();
  var body = {
    model: model,
    max_tokens: 4000,
    output_config: { effort: "low", format: { type: "json_schema", schema: AI_SCHEMA } },
    system: aiSystem_(now),
    messages: [{ role: "user", content: "받은 문자:\n" + text }]
  };
  var headers = { "x-api-key": p.getProperty("ANTHROPIC_API_KEY"), "anthropic-version": "2026-10-21" };
  if (/^claude-(opus|sonnet|fable)/.test(model)) {
    // 안전 필터에 걸려 거절되면 같은 요청을 다른 모델로 자동 재시도
    body.fallbacks = "default";
    headers["anthropic-beta"] = "server-side-fallback-2026-07-01";
  }
  try {
    var res = UrlFetchApp.fetch("https://api.anthropic.com/v1/messages", {
      method: "post", contentType: "application/json", headers: headers, payload: JSON.stringify(body), muteHttpExceptions: true
    });
    if (res.getResponseCode() !== 200) { console.error("AI 오류 " + res.getResponseCode() + ": " + res.getContentText()); return null; }
    var msg = JSON.parse(res.getContentText());
    if (msg.stop_reason !== "end_turn") { console.error("AI 응답 중단: " + msg.stop_reason); return null; }
    var out = msg.content.filter(function (b) { return b.type === "text"; }).map(function (b) { return b.text; }).join("");
    return logic_().fromAI(JSON.parse(out), now);
  } catch (err) {
    console.error("AI 처리 실패: " + err);
    return null;
  }
}

/** 편집기에서 직접 실행해서 AI가 잘 되는지 확인 (결과는 실행 로그에) */
function testAI() {
  if (!aiEnabled_()) { Logger.log("ANTHROPIC_API_KEY 가 없어서 AI가 꺼져 있어요."); return; }
  var r = aiParse_("안녕하세요 다음주 화요일 오후 3시쯤 스탠드 1대랑 벽걸이 2대 청소 될까요? 마포구 공덕동 래미안 101동 이수빈이에요", seoulNow_());
  Logger.log(r ? JSON.stringify(r, null, 2) : "AI 호출 실패 - 위 오류 로그를 확인하세요.");
}


// ===== 휴대폰 화면 (Index.html) =====
var INDEX_HTML = "<!DOCTYPE html>\n<html lang=\"ko\">\n<head>\n<base target=\"_top\">\n<meta charset=\"utf-8\">\n<meta name=\"theme-color\" content=\"#0b6e8a\">\n<link rel=\"preconnect\" href=\"https://fonts.googleapis.com\">\n<link rel=\"preconnect\" href=\"https://fonts.gstatic.com\" crossorigin>\n<link rel=\"stylesheet\" href=\"https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+KR:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap\">\n<style>\n/* 위더스 일정비서 - phone layout. (copied from the demo page)\n   Layout: one phone-width column. Sticky header with two tabs (문자함 / 일정표).\n   문자함 = incoming SMS list -> analysis sheet (verdict, day timeline, alternatives, reply draft). */\n:root {\n  --bg: #eef3f6;\n  --surface: #ffffff;\n  --surface-2: #e4ecf1;\n  --line: #d3dee6;\n  --fg: #13212b;\n  --muted: #5a6b78;\n  --accent: #0b6e8a;\n  --accent-ink: #ffffff;\n  --accent-soft: #d9eff5;\n  --ok: #18774a;\n  --ok-soft: #dcf3e6;\n  --warn: #9a5800;\n  --warn-soft: #fcebd0;\n  --bad: #b42318;\n  --bad-soft: #fde2de;\n  --block: #9fb6c4;\n  --grid: #e7eef2;\n  --c-ac: #d3ecf4;\n  --c-ac-ink: #074b5e;\n  --c-home: #f7e2c6;\n  --c-home-ink: #774100;\n  --font: \"IBM Plex Sans KR\", \"Apple SD Gothic Neo\", \"Malgun Gothic\", system-ui, sans-serif;\n  --mono: \"IBM Plex Mono\", ui-monospace, \"SF Mono\", Menlo, monospace;\n}\n@media (prefers-color-scheme: dark) {\n  :root:not([data-theme=\"light\"]) {\n    --bg: #0d1418; --surface: #152027; --surface-2: #1c2a32; --line: #2a3b45;\n    --fg: #e5eef3; --muted: #92a4b0; --accent: #4fb9d5; --accent-ink: #05252f; --accent-soft: #123540;\n    --ok: #5fcb8f; --ok-soft: #133426; --warn: #e9ab4f; --warn-soft: #3a2a12; --bad: #f2806f; --bad-soft: #3e1b17;\n    --block: #4b6676; --grid: #1e2b33; --c-ac: #123d49; --c-ac-ink: #a3deee; --c-home: #3e2c15; --c-home-ink: #f1c584; color-scheme: dark;\n  }\n}\n:root[data-theme=\"dark\"] {\n  --bg: #0d1418; --surface: #152027; --surface-2: #1c2a32; --line: #2a3b45;\n  --fg: #e5eef3; --muted: #92a4b0; --accent: #4fb9d5; --accent-ink: #05252f; --accent-soft: #123540;\n  --ok: #5fcb8f; --ok-soft: #133426; --warn: #e9ab4f; --warn-soft: #3a2a12; --bad: #f2806f; --bad-soft: #3e1b17;\n  --block: #4b6676; --grid: #1e2b33; --c-ac: #123d49; --c-ac-ink: #a3deee; --c-home: #3e2c15; --c-home-ink: #f1c584; color-scheme: dark;\n}\n* { box-sizing: border-box; }\nbody { background: var(--bg); color: var(--fg); font-family: var(--font); font-size: 15px; line-height: 1.55; -webkit-text-size-adjust: 100%; }\nbody::before { content: \"\"; position: fixed; inset: 0 0 auto; height: env(safe-area-inset-top, 0px); z-index: 10; background: var(--accent); }\nbutton, input, select, textarea { font: inherit; color: inherit; }\nbutton { cursor: pointer; }\n:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }\n.num { font-family: var(--mono); font-variant-numeric: tabular-nums; }\n\nheader.top { position: sticky; top: env(safe-area-inset-top, 0px); z-index: 5; background: var(--accent); color: var(--accent-ink); }\n.top-inner { max-width: 560px; margin: 0 auto; padding-block: 14px 0; padding-left: max(16px, env(safe-area-inset-left, 0px)); padding-right: max(16px, env(safe-area-inset-right, 0px)); }\n.brand { display: flex; align-items: baseline; justify-content: space-between; gap: 12px; flex-wrap: wrap; }\n.brand h1 { margin: 0; font-size: 19px; font-weight: 700; letter-spacing: -0.01em; }\n.brand .today { font-size: 13px; opacity: .85; }\n.tabs { display: flex; gap: 4px; margin-top: 12px; }\n.tab { flex: 1; border: 0; background: transparent; color: var(--accent-ink); opacity: .75; padding: 10px 8px 12px; font-weight: 600; border-bottom: 3px solid transparent; }\n.tab[aria-selected=\"true\"] { opacity: 1; border-bottom-color: var(--accent-ink); }\n.tab .count { font-family: var(--mono); font-size: 12px; margin-left: 4px; }\n\nmain { max-width: 560px; margin: 0 auto; padding-block: 16px 40px; padding-left: max(16px, env(safe-area-inset-left, 0px)); padding-right: max(16px, env(safe-area-inset-right, 0px)); display: flex; flex-direction: column; gap: 20px; }\n.section-label { font-size: 12px; font-weight: 600; letter-spacing: .06em; color: var(--muted); margin: 0 0 8px; }\n.demo-note { font-size: 13px; color: var(--muted); background: var(--surface-2); border-radius: 10px; padding: 10px 12px; margin: 0; }\n\n/* inbox */\n.inbox { display: flex; flex-direction: column; gap: 10px; }\n.sms { text-align: left; width: 100%; background: var(--surface); border: 1px solid var(--line); border-radius: 14px; padding: 12px 14px; display: flex; flex-direction: column; gap: 6px; }\n.sms.is-active { border-color: var(--accent); box-shadow: 0 0 0 2px var(--accent-soft); }\n.sms-top { display: flex; align-items: center; gap: 8px; flex-wrap: wrap; }\n.sms-from { font-weight: 600; font-family: var(--mono); font-size: 14px; }\n.sms-time { font-size: 12px; color: var(--muted); }\n.sms-body { margin: 0; font-size: 14px; color: var(--fg); display: -webkit-box; -webkit-line-clamp: 3; -webkit-box-orient: vertical; overflow: hidden; }\n.sms-cta { font-size: 13px; font-weight: 600; color: var(--accent); }\n.pill { margin-left: auto; font-size: 12px; font-weight: 600; padding: 2px 9px; border-radius: 999px; white-space: nowrap; }\n.pill-new { background: var(--accent-soft); color: var(--accent); }\n.pill-ok { background: var(--ok-soft); color: var(--ok); }\n.pill-bad { background: var(--bad-soft); color: var(--bad); }\n.pill-warn { background: var(--warn-soft); color: var(--warn); }\n\n.paste { background: var(--surface); border: 1px dashed var(--line); border-radius: 14px; padding: 12px; display: flex; flex-direction: column; gap: 8px; }\n.paste label { font-size: 13px; font-weight: 600; }\n.paste textarea { width: 100%; min-height: 84px; resize: vertical; border: 1px solid var(--line); border-radius: 10px; padding: 10px; background: var(--bg); }\n.row { display: flex; gap: 8px; flex-wrap: wrap; }\n.btn { border: 0; border-radius: 10px; padding: 11px 14px; font-weight: 600; font-size: 14px; background: var(--surface-2); color: var(--fg); text-decoration: none; text-align: center; }\n.btn-primary { background: var(--accent); color: var(--accent-ink); }\n.btn-grow { flex: 1; }\n.btn:disabled { opacity: .5; cursor: default; }\n\n/* result sheet */\n.sheet { background: var(--surface); border: 1px solid var(--line); border-radius: 16px; overflow: hidden; }\n.sheet-empty { padding: 28px 18px; text-align: center; color: var(--muted); font-size: 14px; }\n.alert { display: flex; gap: 10px; align-items: flex-start; padding: 12px 14px; background: var(--bad-soft); color: var(--bad); }\n.alert.ok { background: var(--ok-soft); color: var(--ok); }\n.alert.warn { background: var(--warn-soft); color: var(--warn); }\n.alert-icon { flex: none; width: 22px; height: 22px; }\n.alert-title { font-weight: 700; font-size: 15px; }\n.alert-text { font-size: 13px; color: var(--fg); margin-top: 2px; }\n.sheet-body { padding: 14px; display: flex; flex-direction: column; gap: 16px; }\n.quote { margin: 0; font-size: 13px; color: var(--muted); border-left: 3px solid var(--line); padding: 2px 0 2px 10px; }\n.fields { display: grid; grid-template-columns: 88px 1fr; gap: 8px 12px; align-items: center; margin: 0; }\n.fields dt { font-size: 13px; color: var(--muted); }\n.fields dd { margin: 0; min-width: 0; font-size: 14px; overflow-wrap: anywhere; }\n.fields input, .fields select { width: 100%; max-width: 220px; border: 1px solid var(--line); background: var(--bg); border-radius: 8px; padding: 6px 8px; font-size: 14px; }\n.hint { font-size: 12px; color: var(--muted); }\n.missing { color: var(--warn); }\n\n.conflict-list { display: flex; flex-direction: column; gap: 6px; margin: 0; padding: 0; list-style: none; }\n.conflict-list li { font-size: 14px; background: var(--bad-soft); border-radius: 8px; padding: 8px 10px; }\n\n/* week calendar */\n.wk { --hr: 44px; background: var(--surface); border: 1px solid var(--line); border-radius: 14px; overflow: hidden; }\n.wk-head, .wk-body { display: grid; }\n.wk-head { border-bottom: 1px solid var(--line); }\n.wk-dh { text-align: center; padding: 6px 0 4px; font-size: 12px; color: var(--muted); min-width: 0; }\n.wk-dh b { display: block; width: 28px; height: 28px; line-height: 28px; margin: 1px auto 0; border-radius: 999px; font-family: var(--mono); font-size: 15px; font-weight: 600; color: var(--fg); }\n.wk-dh.is-today b { background: var(--accent); color: var(--accent-ink); }\n.wk-dh.is-past { opacity: .55; }\n.wk-n { display: block; font-size: 10px; font-weight: 600; color: var(--accent); }\n.wk-body { padding-top: 6px; padding-bottom: 6px; }\n.wk-axis { position: relative; height: calc(var(--hr) * var(--hours, 11)); }\n.wk-axis span { position: absolute; right: 4px; transform: translateY(-50%); font-family: var(--mono); font-size: 10px; color: var(--muted); }\n.wk-col { position: relative; height: calc(var(--hr) * var(--hours, 11)); border-left: 1px solid var(--line); min-width: 0;\n  background-image: repeating-linear-gradient(to bottom, var(--grid) 0, var(--grid) 1px, transparent 1px, transparent var(--hr)); }\n.wk-col.is-past, .wk-col.is-off { background-color: var(--surface-2); }\n.wk-col.pickable { cursor: pointer; }\n.ev { position: absolute; left: 2px; right: 2px; z-index: 1; display: flex; flex-direction: column; gap: 0; overflow: hidden; text-align: left;\n  border: 0; border-left: 3px solid var(--accent); border-radius: 5px; padding: 2px 2px 2px 3px; font-size: 11px; letter-spacing: -0.02em; line-height: 1.25;\n  background: var(--c-ac); color: var(--c-ac-ink); }\n.ev.home { background: var(--c-home); color: var(--c-home-ink); border-left-color: var(--c-home-ink); }\n.ev-t { font-family: var(--mono); font-size: 10px; opacity: .85; }\n.ev-n { font-weight: 700; white-space: nowrap; overflow: hidden; }\n.ev-l { overflow: hidden; word-break: keep-all; }\n.sheet .wk { margin: 0 -14px; border-radius: 0; border-left: 0; border-right: 0; }\n.ev.is-sel { box-shadow: 0 0 0 2px var(--fg); }\n.ev.hit { box-shadow: 0 0 0 2px var(--bad); }\n.ev.prop { z-index: 2; pointer-events: none; border: 2px dashed var(--ok); background: color-mix(in srgb, var(--ok) 20%, var(--surface)); color: var(--ok); }\n.ev.prop.bad { border-color: var(--bad); background: color-mix(in srgb, var(--bad) 18%, var(--surface)); color: var(--bad); }\n.wk-now { position: absolute; left: 0; right: 0; height: 2px; background: var(--bad); z-index: 3; pointer-events: none; }\n.legend { display: flex; gap: 6px 14px; font-size: 12px; color: var(--muted); flex-wrap: wrap; }\n.legend i { display: inline-block; width: 12px; height: 10px; border-radius: 2px; margin-right: 5px; vertical-align: -1px; }\n\n.alts { display: flex; flex-direction: column; gap: 8px; }\n.alt { display: flex; align-items: center; justify-content: space-between; gap: 10px; width: 100%; text-align: left; border: 1px solid var(--line); background: var(--surface); border-radius: 10px; padding: 10px 12px; }\n.alt[aria-pressed=\"true\"] { border-color: var(--ok); background: var(--ok-soft); }\n.alt-date { font-weight: 600; }\n.alt-why { font-size: 12px; color: var(--muted); }\n.alt-mark { flex: none; width: 20px; height: 20px; border-radius: 50%; border: 2px solid var(--line); }\n.alt[aria-pressed=\"true\"] .alt-mark { border-color: var(--ok); background: var(--ok); box-shadow: inset 0 0 0 3px var(--ok-soft); }\n\n.reply { width: 100%; min-height: 128px; border: 1px solid var(--line); border-radius: 10px; padding: 10px; background: var(--bg); font-size: 14px; resize: vertical; }\n.phone-line { font-size: 13px; color: var(--muted); }\n.done-banner { background: var(--ok-soft); color: var(--ok); border-radius: 10px; padding: 10px 12px; font-weight: 600; font-size: 14px; }\n\n/* schedule */\n.sched-summary { display: grid; grid-template-columns: repeat(3, 1fr); gap: 8px; }\n.stat { background: var(--surface); border: 1px solid var(--line); border-radius: 12px; padding: 10px 12px; }\n.stat b { display: block; font-family: var(--mono); font-size: 22px; font-weight: 600; }\n.stat span { font-size: 12px; color: var(--muted); }\n/* schedule tab */\n.wk-nav { display: flex; align-items: center; gap: 8px; }\n.wk-range { flex: 1; text-align: center; font-weight: 700; font-size: 16px; }\n.wk-range small { display: block; font-size: 12px; font-weight: 500; color: var(--muted); }\n.wk-btn { flex: none; width: 44px; height: 40px; border: 1px solid var(--line); background: var(--surface); border-radius: 10px; font-size: 20px; line-height: 1; }\n.detail { background: var(--surface); border: 1px solid var(--line); border-radius: 14px; padding: 12px 14px; display: flex; flex-direction: column; gap: 4px; font-size: 14px; }\n.detail-top { display: flex; justify-content: space-between; align-items: center; gap: 8px; }\n.detail-when { font-weight: 700; }\n.job-who { font-size: 13px; color: var(--muted); overflow-wrap: anywhere; }\n.job-del { border: 1px solid var(--line); background: transparent; color: var(--muted); font-size: 13px; padding: 5px 10px; border-radius: 8px; }\n.job-del.armed { color: var(--bad); background: var(--bad-soft); border-color: transparent; font-weight: 600; }\n.foot { display: flex; justify-content: center; gap: 8px; flex-wrap: wrap; }\n\n.howto { background: var(--surface); border: 1px solid var(--line); border-radius: 14px; padding: 14px; }\n.howto h2 { font-size: 16px; margin: 0 0 10px; }\n.howto ol { margin: 0; padding-left: 20px; display: flex; flex-direction: column; gap: 8px; font-size: 14px; }\n.howto li span { color: var(--muted); font-size: 13px; display: block; }\n\n.toast { position: fixed; left: 50%; bottom: calc(20px + env(safe-area-inset-bottom, 0px)); transform: translateX(-50%); background: var(--fg); color: var(--bg); padding: 10px 16px; border-radius: 999px; font-size: 14px; font-weight: 600; z-index: 20; max-width: calc(100% - 32px); }\n@media (prefers-reduced-motion: no-preference) {\n  .sheet.fresh { animation: rise .25s ease-out; }\n  @keyframes rise { from { transform: translateY(6px); opacity: .6; } to { transform: none; opacity: 1; } }\n}\n\n/* app-only additions */\nhtml { -webkit-tap-highlight-color: transparent; }\n:root { padding-top: env(safe-area-inset-top, 0px); padding-bottom: env(safe-area-inset-bottom, 0px); }\nbody { margin: 0; }\n.refresh { border: 1px solid color-mix(in srgb, var(--accent-ink) 45%, transparent); background: transparent; color: var(--accent-ink); border-radius: 999px; padding: 4px 12px; font-size: 13px; font-weight: 600; }\n.sms.is-done { opacity: .7; }\n.fields input[type=\"text\"], .fields input[type=\"tel\"] { max-width: 100%; }\n.howto .btn { font-size: 13px; padding: 8px 12px; }\n.phone-line a { color: var(--accent); font-weight: 600; }\n.detail a { color: var(--accent); }\n.hook { display: block; margin-top: 6px; padding: 8px 10px; border-radius: 8px; background: var(--surface-2); color: var(--fg); font-size: 12px; overflow-wrap: anywhere; user-select: all; }\n</style>\n</head>\n<body>\n<header class=\"top\">\n  <div class=\"top-inner\">\n    <div class=\"brand\">\n      <h1>위더스 일정비서</h1>\n      <button type=\"button\" class=\"refresh\" id=\"refreshBtn\" aria-label=\"새로고침\">새로고침</button>\n    </div>\n    <div class=\"tabs\" role=\"tablist\">\n      <button class=\"tab\" role=\"tab\" id=\"tab-inbox\" aria-selected=\"true\" data-view=\"inbox\">문의<span class=\"count\" id=\"inboxCount\"></span></button>\n      <button class=\"tab\" role=\"tab\" id=\"tab-sched\" aria-selected=\"false\" data-view=\"sched\">일정표<span class=\"count\" id=\"jobCount\"></span></button>\n    </div>\n  </div>\n</header>\n\n<main>\n  <section id=\"view-inbox\" class=\"view\">\n    <div style=\"display:flex;flex-direction:column;gap:20px\">\n      <div>\n        <p class=\"section-label\">들어온 문의 <span id=\"todayLabel\" style=\"font-weight:500;letter-spacing:0\"></span></p>\n        <div class=\"inbox\" id=\"inbox\"></div>\n      </div>\n\n      <form class=\"paste\" id=\"pasteForm\">\n        <label for=\"pasteText\">전화로 받은 예약이나 문자 직접 넣기</label>\n        <textarea id=\"pasteText\" placeholder=\"예) 12월 3일 오후 2시 벽걸이 2대 김철수 010-1234-5678 마포구 공덕동 래미안 101동\"></textarea>\n        <div class=\"row\"><button class=\"btn btn-primary btn-grow\" type=\"submit\" id=\"pasteBtn\">일정 확인하기</button></div>\n      </form>\n\n      <div id=\"resultAnchor\">\n        <p class=\"section-label\">확인 결과</p>\n        <div class=\"sheet\" id=\"result\"><div class=\"sheet-empty\">문의를 누르면 날짜·시간·작업을 읽고<br>일정표와 겹치는지 알려줘요.</div></div>\n      </div>\n\n      <div class=\"howto\" id=\"settings\"></div>\n    </div>\n  </section>\n\n  <section id=\"view-sched\" class=\"view\" hidden>\n    <div style=\"display:flex;flex-direction:column;gap:14px\">\n      <div class=\"wk-nav\">\n        <button type=\"button\" class=\"wk-btn\" id=\"wkPrev\" aria-label=\"지난 주\">‹</button>\n        <div class=\"wk-range\"><span id=\"wkRange\"></span><small id=\"wkSub\"></small></div>\n        <button type=\"button\" class=\"wk-btn\" id=\"wkNext\" aria-label=\"다음 주\">›</button>\n      </div>\n      <div class=\"sched-summary\" id=\"schedSummary\"></div>\n      <div id=\"wkGrid\"></div>\n      <div class=\"legend\"><span><i style=\"background:var(--c-ac)\"></i>에어컨 청소</span><span><i style=\"background:var(--c-home)\"></i>집 청소</span><span><i style=\"background:var(--bad);height:2px;vertical-align:3px\"></i>지금</span><span>휴무일은 일정이 있을 때만 보여요</span></div>\n      <div id=\"wkDetail\"></div>\n      <div class=\"foot\"><button class=\"btn\" id=\"wkToday\" type=\"button\">이번 주로</button></div>\n    </div>\n  </section>\n</main>\n\n<div class=\"toast\" id=\"toast\" hidden></div>\n\n<script>\nvar KEY = \"<?= key ?>\";\nvar APP_URL = \"<?= appUrl ?>\";\nvar CFG = <?!= cfgJson ?>;\nvar L = (<?!= logicSrc ?>)(CFG);\nvar state = <?!= bootJson ?>;\n</script>\n<script>\n(function () {\n  \"use strict\";\n  var DOW = L.DOW, OPEN = L.OPEN, CLOSE = L.CLOSE, FULL_START = L.FULL_START, FULL_DUR = L.FULL_DUR, HR = 44;\n  var key = L.key, fromKey = L.fromKey, addDays = L.addDays, hm = L.hm, fmtDate = L.fmtDate, fmtTime = L.fmtTime, fmtDur = L.fmtDur, whenText = L.whenText, shortLabel = L.shortLabel;\n  function now() { return new Date(); }\n  function today0() { return L.day0(now()); }\n  function tkey() { return key(today0()); }\n  function dayJobs(k) { return L.jobsOn(state.jobs, k); }\n  function esc(s) { return String(s == null ? \"\" : s).replace(/[&<>\"']/g, function (c) { return { \"&\": \"&amp;\", \"<\": \"&lt;\", \">\": \"&gt;\", '\"': \"&quot;\", \"'\": \"&#39;\" }[c]; }); }\n  var $ = function (id) { return document.getElementById(id); };\n\n  // ---------- 서버 호출 ----------\n  var busy = false;\n  function run(fn) {\n    var args = [KEY].concat([].slice.call(arguments, 1));\n    return new Promise(function (resolve, reject) {\n      google.script.run\n        .withSuccessHandler(resolve)\n        .withFailureHandler(function (e) { toast((e && e.message) || \"연결이 안 돼요. 잠시 뒤 다시 해 주세요.\"); reject(e); })[fn].apply(null, args);\n    });\n  }\n  function toast(msg) {\n    var t = $(\"toast\"); t.textContent = msg; t.hidden = false;\n    clearTimeout(toast._t); toast._t = setTimeout(function () { t.hidden = true; }, 2400);\n  }\n\n  // ---------- 문의 목록 ----------\n  var cur = null, ev = null, sel = null, added = false, armedDel = null, armedDismiss = false;\n\n  function itemPill(m) {\n    if (m.status === \"done\") return { tone: \"ok\", text: \"일정 등록됨\" };\n    if (!m.parsed) return { tone: \"warn\", text: \"확인 필요\" };\n    return L.pill(L.evaluate(m.parsed, state.jobs, now()).status);\n  }\n  function renderInbox() {\n    var items = state.inbox;\n    $(\"inbox\").innerHTML = items.length ? items.map(function (m) {\n      var p = itemPill(m);\n      return '<button type=\"button\" class=\"sms' + (cur && cur.msgId === m.id ? \" is-active\" : \"\") + (m.status === \"done\" ? \" is-done\" : \"\") + '\" data-id=\"' + esc(m.id) + '\">' +\n        '<div class=\"sms-top\"><span class=\"sms-from\">' + esc(m.source === \"manual\" ? \"직접 입력\" : (m.from || \"번호 없음\")) + '</span><span class=\"sms-time\">' + esc(String(m.receivedAt).slice(5).replace(\"-\", \"/\")) + \"</span>\" +\n        '<span class=\"pill pill-' + p.tone + '\">' + esc(p.text) + \"</span></div>\" +\n        '<p class=\"sms-body\">' + esc(m.text) + \"</p></button>\";\n    }).join(\"\") : '<div class=\"sheet-empty\" style=\"background:var(--surface);border:1px solid var(--line);border-radius:14px\">아직 들어온 문의가 없어요.<br>문자 전달 앱이 연결되면 손님 문자가 여기에 자동으로 떠요.</div>';\n    var pending = items.filter(function (m) { return m.status === \"new\"; }).length;\n    $(\"inboxCount\").textContent = pending ? \" \" + pending : \"\";\n  }\n\n  function renderSettings() {\n    var s = state.settings, h = \"<h2>연결 상태</h2><ol>\", hook = APP_URL + \"?key=\" + KEY;\n    h += \"<li>문자 전달 앱 \" + (s.lastPostAt ? \"연결됨<span>마지막으로 문자를 받은 시각: \" + esc(s.lastPostAt) + \"</span>\" : \"아직 연결 기록 없음<span>전달 앱의 Webhook URL 칸에 아래 주소를 넣고, 다른 폰에서 테스트 문자를 보내 보세요.</span>\") +\n      '<span class=\"hook num\" id=\"hookUrl\">' + esc(hook) + '</span><button type=\"button\" class=\"btn\" id=\"copyHook\" style=\"margin-top:6px\">전달 앱 주소 복사</button></li>';\n    if (s.kakao) h += '<li>카카오톡 알림 켜짐<span>새 문의가 오면 카카오톡 \\'나와의 채팅\\'으로 알려줘요.</span><button type=\"button\" class=\"btn\" id=\"testKakao\" style=\"margin-top:6px\">테스트 알림 보내기</button></li>';\n    else if (s.kakaoReady) h += '<li>카카오톡 알림이 아직 연결 안 됐어요<span>아빠 카카오 계정으로 한 번만 로그인하면 돼요.</span><a class=\"btn btn-primary\" style=\"margin-top:6px;display:inline-block\" href=\"' + esc(APP_URL + \"?key=\" + KEY + \"&kakao=login\") + '\" target=\"_top\">카카오톡 연결하기</a></li>';\n    else h += \"<li>카카오톡 알림 꺼짐<span>설명서의 '카카오톡 알림' 단계를 하면 켜져요.</span></li>\";\n    h += s.ai ? \"<li>AI로 문자 읽는 중<span>\" + esc(s.aiModel) + \" · 문자 한 통마다 사용료가 조금 들어요.</span></li>\"\n      : \"<li>무료 규칙으로 문자 읽는 중<span>사용료 0원. 날짜를 잘 못 읽으면 결과 화면에서 고칠 수 있어요.</span></li>\";\n    $(\"settings\").innerHTML = h + \"</ol>\";\n  }\n\n  // ---------- 주간 달력 ----------\n  function weekMon(k) { return key(L.mondayOf(fromKey(k))); }\n  function weekRange(mk) {\n    var a = fromKey(mk), b = addDays(a, 5);\n    return (a.getMonth() + 1) + \"월 \" + a.getDate() + \"일 – \" + (a.getMonth() === b.getMonth() ? \"\" : (b.getMonth() + 1) + \"월 \") + b.getDate() + \"일\";\n  }\n  function weekGrid(mk, o) {\n    o = o || {};\n    var mon = fromKey(mk), tk = tkey(), days = [];\n    for (var i = 0; i < 7; i++) days.push(key(addDays(mon, i)));\n    days = days.filter(function (k) { return !L.isClosed(k) || dayJobs(k).length || (o.prop && o.prop.date === k); });\n    var cols = \"26px repeat(\" + days.length + \", minmax(0, 1fr))\";\n    var y = function (m) { return (Math.max(OPEN, Math.min(CLOSE, m)) - OPEN) / 60 * HR; };\n    var hgt = function (s, d) { return Math.max(18, y(s + d) - y(s)).toFixed(1); };\n    var hours = (CLOSE - OPEN) / 60;\n    var h = '<div class=\"wk\" style=\"--hours:' + hours + '\"><div class=\"wk-head\" style=\"grid-template-columns:' + cols + '\"><span></span>';\n    days.forEach(function (k) {\n      var d = fromKey(k), n = dayJobs(k).length;\n      h += '<div class=\"wk-dh' + (k === tk ? \" is-today\" : \"\") + (k < tk ? \" is-past\" : \"\") + '\">' + DOW[d.getDay()] + \"<b>\" + d.getDate() + '</b><span class=\"wk-n\">' + (n ? n + \"건\" : \"&nbsp;\") + \"</span></div>\";\n    });\n    h += '</div><div class=\"wk-body\" style=\"grid-template-columns:' + cols + '\"><div class=\"wk-axis\" aria-hidden=\"true\">';\n    for (var t = OPEN; t <= CLOSE; t += 60) h += '<span style=\"top:' + y(t) + 'px\">' + (t / 60) + \"</span>\";\n    h += \"</div>\";\n    days.forEach(function (k) {\n      h += '<div class=\"wk-col' + (k < tk ? \" is-past\" : \"\") + (L.isClosed(k) ? \" is-off\" : \"\") + (o.pick ? \" pickable\" : \"\") + '\" data-date=\"' + k + '\">';\n      dayJobs(k).forEach(function (j) {\n        var s = j.fullDay ? FULL_START : j.start, d = j.fullDay ? FULL_DUR : j.dur;\n        h += '<button type=\"button\" class=\"ev' + (/^에어컨/.test(j.label) ? \"\" : \" home\") + ((o.hits || []).indexOf(j.id) >= 0 ? \" hit\" : \"\") + (o.selId === j.id ? \" is-sel\" : \"\") +\n          '\" data-job=\"' + esc(j.id) + '\" style=\"top:' + y(s) + \"px;height:\" + hgt(s, d) + 'px\" aria-label=\"' + esc(fmtDate(k) + \" \" + fmtTime(s) + \" \" + (j.customer || \"\") + \" \" + j.label) + '\">' +\n          '<span class=\"ev-t\">' + (j.fullDay ? \"종일\" : hm(s)) + '</span><span class=\"ev-n\">' + esc(j.customer || \"고객\") + '</span><span class=\"ev-l\">' + esc(shortLabel(j.label)) + \"</span></button>\";\n      });\n      if (o.prop && o.prop.date === k) {\n        var p = o.prop;\n        h += '<div class=\"ev prop' + (p.bad ? \" bad\" : \"\") + '\" style=\"top:' + y(p.start) + \"px;height:\" + hgt(p.start, p.dur) + 'px\"><span class=\"ev-t\">' + (p.full ? \"종일\" : hm(p.start)) + '</span><span class=\"ev-n\">' + (p.bad ? \"요청\" : \"새 예약\") + \"</span></div>\";\n      }\n      if (k === tk) { var nm = L.minsOf(now()); if (nm > OPEN && nm < CLOSE) h += '<span class=\"wk-now\" style=\"top:' + y(nm) + 'px\"></span>'; }\n      h += \"</div>\";\n    });\n    return h + \"</div></div>\";\n  }\n\n  // ---------- 확인 결과 ----------\n  var ICON_BAD = '<svg class=\"alert-icon\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" stroke-linecap=\"round\"><path d=\"M12 3 2 20h20L12 3z\"/><path d=\"M12 10v4M12 17.5v.01\"/></svg>';\n  var ICON_OK = '<svg class=\"alert-icon\" viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\" stroke-linecap=\"round\" stroke-linejoin=\"round\"><circle cx=\"12\" cy=\"12\" r=\"9\"/><path d=\"m8 12 3 3 5-6\"/></svg>';\n  var EMPTY_RESULT = '<div class=\"sheet-empty\">문의를 누르면 날짜·시간·작업을 읽고<br>일정표와 겹치는지 알려줘요.</div>';\n\n  function renderResult(fresh) {\n    var box = $(\"result\");\n    if (!cur) { box.innerHTML = EMPTY_RESULT; return; }\n    var bad = ev.status !== \"ok\" && ev.status !== \"suggest\", h = \"\", title, text, cls;\n    if (added) { cls = \"ok\"; title = \"일정표에 넣었어요\"; text = whenText(sel.date, sel.start, cur.full) + \" · \" + esc(cur.label); }\n    else if (ev.status === \"ok\") { cls = \"ok\"; title = \"이 시간 비어 있어요\"; text = whenText(cur.date, ev.start, cur.full) + \"에 넣을 수 있어요.\"; }\n    else if (ev.status === \"suggest\") { cls = \"ok\"; title = \"빈 시간을 찾았어요\"; text = \"고객 요청(\" + esc(cur.timeText || \"시간 무관\") + \")에 맞춰 \" + fmtDate(cur.date) + \" \" + fmtTime(ev.start) + \" 시작을 추천해요.\"; }\n    else if (ev.status === \"conflict\") {\n      cls = \"\"; title = \"일정이 겹쳐요\";\n      text = ev.windowFull ? fmtDate(cur.date) + \" \" + esc(cur.timeText) + \"에는 넣을 자리가 없어요.\" : whenText(cur.date, cur.start, cur.full) + \" 요청이 기존 예약 \" + ev.conflicts.length + \"건과 겹쳐요.\";\n    }\n    else if (ev.status === \"closed\") { cls = \"warn\"; title = \"휴무일이에요\"; text = fmtDate(cur.date) + \" 요청이에요. 다른 날을 골라 주세요.\"; }\n    else if (ev.status === \"outside\") { cls = \"warn\"; title = \"작업 시간 밖이에요\"; text = fmtTime(cur.start) + \" 시작이면 \" + fmtTime(cur.start + cur.dur) + \"에 끝나요. 작업은 \" + fmtTime(OPEN) + \"~\" + fmtTime(CLOSE) + \"예요.\"; }\n    else if (ev.status === \"past\") { cls = \"warn\"; title = \"이미 지난 시간이에요\"; text = \"날짜를 다시 확인해 주세요.\"; }\n    else { cls = \"warn\"; title = \"날짜를 못 찾았어요\"; text = \"아래에서 날짜를 직접 골라 주세요.\"; }\n    h += '<div class=\"alert ' + cls + '\" role=\"status\">' + (cls === \"ok\" ? ICON_OK : ICON_BAD) + '<div><div class=\"alert-title\">' + title + '</div><div class=\"alert-text\">' + text + \"</div></div></div>\";\n    h += '<div class=\"sheet-body\"><blockquote class=\"quote\">' + esc(cur.text) + \"</blockquote>\";\n\n    var durOpts = [60, 90, 120, 150, 180, 240, 300, 360];\n    if (!cur.full && durOpts.indexOf(cur.dur) < 0) { durOpts.push(cur.dur); durOpts.sort(function (a, b) { return a - b; }); }\n    var dis = added ? \" disabled\" : \"\";\n    h += '<div><p class=\"section-label\">문자에서 읽은 내용' + (cur.via === \"ai\" ? \" · AI\" : \"\") + '</p><dl class=\"fields\">';\n    h += '<dt><label for=\"fDate\">날짜</label></dt><dd><input type=\"date\" id=\"fDate\" value=\"' + (cur.date || \"\") + '\"' + dis + \">\" + (cur.dateText ? ' <span class=\"hint\">“' + esc(cur.dateText) + '”</span>' : \"\") + \"</dd>\";\n    h += '<dt><label for=\"fTime\">시간</label></dt><dd>' + (cur.full ? \"하루 종일\" :\n      '<input type=\"time\" id=\"fTime\" step=\"1800\" value=\"' + (cur.start != null ? hm(cur.start) : \"\") + '\"' + dis + \">\" + (cur.start == null ? ' <span class=\"hint\">' + (cur.timeText ? \"“\" + esc(cur.timeText) + \"” → 빈 시간 자동 선택\" : \"시간 언급 없음 → 빈 시간 자동 선택\") + \"</span>\" : \"\")) + \"</dd>\";\n    h += '<dt><label for=\"fLabel\">작업</label></dt><dd><input type=\"text\" id=\"fLabel\" value=\"' + esc(cur.label) + '\"' + dis + \"></dd>\";\n    h += '<dt><label for=\"fDur\">예상 소요</label></dt><dd>' + (cur.full ? \"하루 종일\" :\n      '<select id=\"fDur\"' + dis + \">\" + durOpts.map(function (d) { return '<option value=\"' + d + '\"' + (d === cur.dur ? \" selected\" : \"\") + \">\" + fmtDur(d) + \"</option>\"; }).join(\"\") + \"</select>\") + \"</dd>\";\n    h += '<dt><label for=\"fName\">고객</label></dt><dd><input type=\"text\" id=\"fName\" value=\"' + esc(cur.customer) + '\" placeholder=\"이름 없음\"' + dis + \"></dd>\";\n    h += '<dt><label for=\"fPhone\">연락처</label></dt><dd><input type=\"tel\" id=\"fPhone\" value=\"' + esc(cur.phone) + '\"' + dis + \"></dd>\";\n    h += '<dt><label for=\"fAddr\">주소</label></dt><dd><input type=\"text\" id=\"fAddr\" value=\"' + esc(cur.address) + '\" placeholder=\"주소 없음 · 답장에서 물어보세요\"' + dis + \"></dd>\";\n    h += \"</dl></div>\";\n\n    if (cur.date) {\n      var pdur = cur.full ? FULL_DUR : cur.dur;\n      var prop = sel ? { date: sel.date, start: sel.start, dur: pdur, full: cur.full }\n        : (cur.start != null && ev.status !== \"nodate\" ? { date: cur.date, start: cur.start, dur: pdur, full: cur.full, bad: true } : null);\n      var focus = sel ? sel.date : cur.date;\n      h += '<div><p class=\"section-label\">' + weekRange(weekMon(focus)) + \" 주간 일정</p>\" +\n        weekGrid(weekMon(focus), { prop: added ? null : prop, hits: added ? [] : ev.conflicts.map(function (j) { return j.id; }), pick: !added });\n      h += '<div class=\"legend\" style=\"margin-top:8px\"><span><i style=\"background:var(--c-ac)\"></i>기존 예약</span>' + (prop && !added ? '<span><i style=\"background:' + (prop.bad ? \"var(--bad)\" : \"var(--ok)\") + '\"></i>' + (prop.bad ? \"요청한 시간\" : \"넣을 시간\") + \"</span>\" : \"\") + \"<span>이동시간 \" + L.BUF + \"분 포함</span></div>\";\n      if (!added) h += '<p class=\"hint\" style=\"margin:4px 0 0\">달력의 빈 칸을 누르면 그 시간으로 다시 확인해요.</p>';\n      if (ev.conflicts.length && !added) {\n        h += '<ul class=\"conflict-list\" style=\"margin-top:10px\">' + ev.conflicts.map(function (j) {\n          return '<li><b class=\"num\">' + (j.fullDay ? \"종일\" : hm(j.start) + \"–\" + hm(j.start + j.dur)) + \"</b> \" + esc(j.label) + '<br><span class=\"hint\">' + esc(j.customer) + \" · \" + esc(j.address) + \"</span></li>\";\n        }).join(\"\") + \"</ul>\";\n      }\n      h += \"</div>\";\n    }\n\n    if (bad && ev.alts.length && !added) {\n      h += '<div><p class=\"section-label\">대신 가능한 시간</p><div class=\"alts\">' + ev.alts.map(function (a, i) {\n        var on = sel && sel.date === a.date && sel.start === a.start;\n        var why = a.sameDay ? \"같은 날 다른 시간\" : (cur.full ? \"그날 예약 없음\" : (cur.start != null && a.start === cur.start ? \"같은 시간, 다른 날\" : \"가장 가까운 빈 시간\"));\n        return '<button type=\"button\" class=\"alt\" data-alt=\"' + i + '\" aria-pressed=\"' + (on ? \"true\" : \"false\") + '\"><span><span class=\"alt-date\">' + whenText(a.date, a.start, cur.full) + '</span><br><span class=\"alt-why\">' + why + '</span></span><span class=\"alt-mark\"></span></button>';\n      }).join(\"\") + \"</div></div>\";\n    }\n\n    if (!added && sel) h += '<button type=\"button\" class=\"btn btn-primary\" id=\"addBtn\">' + whenText(sel.date, sel.start, cur.full) + \" 일정표에 넣기</button>\";\n    else if (!added && bad && ev.alts.length) h += '<p class=\"hint\" style=\"margin:0\">위에서 시간을 고르면 일정표에 바로 넣을 수 있어요. 고객에게 먼저 물어보려면 아래 답장을 보내세요.</p>';\n\n    var reply = L.replyText(cur, ev, sel, added);\n    h += '<div><p class=\"section-label\"><label for=\"replyText\">고객에게 보낼 답장</label></p><textarea class=\"reply\" id=\"replyText\">' + esc(reply) + \"</textarea>\";\n    h += '<div class=\"row\" style=\"margin-top:8px\"><button type=\"button\" class=\"btn btn-grow\" id=\"copyBtn\">답장 복사</button>';\n    if (cur.phone) h += '<a class=\"btn btn-grow\" id=\"smsLink\" target=\"_top\" href=\"sms:' + esc(cur.phone) + \"?body=\" + encodeURIComponent(reply) + '\">문자앱 열기</a>';\n    h += \"</div>\" + (cur.phone ? '<p class=\"phone-line\">받는 사람 <span class=\"num\">' + esc(cur.phone) + '</span> · <a href=\"tel:' + esc(cur.phone) + '\" target=\"_top\">전화 걸기</a></p>' : \"\") + \"</div>\";\n    h += '<button type=\"button\" class=\"btn\" id=\"dismissBtn\" style=\"align-self:center\">' + (armedDismiss ? \"정말 목록에서 지울까요?\" : \"목록에서 지우기\") + \"</button>\";\n    h += \"</div>\";\n    box.innerHTML = h;\n    box.classList.remove(\"fresh\"); if (fresh) { void box.offsetWidth; box.classList.add(\"fresh\"); }\n  }\n\n  function reevaluate(fresh) {\n    ev = L.evaluate(cur, state.jobs, now());\n    if (!added) sel = (ev.status === \"ok\" || ev.status === \"suggest\") ? { date: cur.date, start: ev.start } : null;\n    renderInbox(); renderResult(fresh);\n  }\n\n  function openItem(id, scroll) {\n    var m = state.inbox.filter(function (x) { return x.id === id; })[0];\n    if (!m) return;\n    var p = m.parsed || L.parse(m.text, now());\n    cur = JSON.parse(JSON.stringify(p)); cur.text = m.text; cur.msgId = m.id;\n    if (!cur.phone && /^\\+?\\d[\\d-]{7,}$/.test(m.from || \"\")) cur.phone = m.from;\n    sel = null; added = m.status === \"done\"; armedDismiss = false;\n    if (added) {\n      var j = state.jobs.filter(function (x) { return x.inboxId === m.id; })[0];\n      sel = j ? { date: j.date, start: j.start } : (cur.date ? { date: cur.date, start: cur.start || 540 } : null);\n      if (!sel) added = false; else { cur.date = sel.date; if (!cur.full) cur.start = sel.start; }\n    }\n    reevaluate(true);\n    if (scroll) $(\"resultAnchor\").scrollIntoView({ behavior: \"smooth\", block: \"start\" });\n  }\n\n  // ---------- 일정표 탭 ----------\n  var viewMon = null, selJob = null;\n  function defaultMon() { var T = today0(), m = L.mondayOf(T); if (T.getDay() === 0) m = addDays(m, 7); return key(m); }\n  function renderSched() {\n    if (!viewMon) viewMon = defaultMon();\n    var mon = fromKey(viewMon), endK = key(addDays(mon, 6)), tk = tkey();\n    var wj = state.jobs.filter(function (j) { return j.date >= viewMon && j.date <= endK; });\n    var busyMin = wj.reduce(function (s, j) { return s + (j.fullDay ? FULL_DUR : j.dur); }, 0), free = 0;\n    for (var i = 0; i < 7; i++) { var k = key(addDays(mon, i)); if (k >= tk && !L.isClosed(k) && !dayJobs(k).length) free++; }\n    var diff = Math.round((mon - fromKey(defaultMon())) / 6048e5);\n    $(\"wkRange\").textContent = weekRange(viewMon);\n    $(\"wkSub\").textContent = diff === 0 ? \"이번 주\" : diff === 1 ? \"다음 주\" : diff === -1 ? \"지난 주\" : diff > 0 ? diff + \"주 뒤\" : -diff + \"주 전\";\n    $(\"schedSummary\").innerHTML =\n      '<div class=\"stat\"><b>' + wj.length + \"</b><span>예약</span></div>\" +\n      '<div class=\"stat\"><b>' + (busyMin ? (busyMin / 60).toFixed(busyMin % 60 ? 1 : 0) : 0) + \"</b><span>작업 시간</span></div>\" +\n      '<div class=\"stat\"><b>' + free + \"</b><span>남은 빈 날</span></div>\";\n    $(\"wkGrid\").innerHTML = weekGrid(viewMon, { selId: selJob });\n    var j = state.jobs.filter(function (x) { return x.id === selJob; })[0];\n    $(\"wkDetail\").innerHTML = j ?\n      '<div class=\"detail\"><div class=\"detail-top\"><span class=\"detail-when\">' + fmtDate(j.date) + \" · \" + (j.fullDay ? \"하루 종일\" : '<span class=\"num\">' + hm(j.start) + \"–\" + hm(j.start + j.dur) + \"</span>\") + \"</span>\" +\n      '<button type=\"button\" class=\"job-del' + (armedDel === j.id ? \" armed\" : \"\") + '\" data-del=\"' + esc(j.id) + '\">' + (armedDel === j.id ? \"정말 삭제\" : \"삭제\") + \"</button></div>\" +\n      \"<div>\" + esc(j.label) + '</div><div class=\"job-who\">' + esc(j.customer || \"이름 없음\") +\n      (j.phone ? ' · <a class=\"num\" href=\"tel:' + esc(j.phone) + '\" target=\"_top\">' + esc(j.phone) + \"</a>\" : \"\") + (j.address ? \"<br>\" + esc(j.address) : \"\") + \"</div></div>\"\n      : '<p class=\"hint\" style=\"margin:0;text-align:center\">일정을 누르면 고객 정보가 여기 나와요.</p>';\n    $(\"jobCount\").textContent = \" \" + state.jobs.filter(function (x) { return x.date >= tk; }).length;\n  }\n\n  function applyState(s) {\n    state = s;\n    if (cur && !state.inbox.some(function (m) { return m.id === cur.msgId; })) cur = null;\n    renderInbox(); renderSched(); renderSettings();\n    if (cur) reevaluate(false); else renderResult(false);\n  }\n\n  // ---------- 이벤트 ----------\n  function showView(v) {\n    $(\"view-inbox\").hidden = v !== \"inbox\"; $(\"view-sched\").hidden = v !== \"sched\";\n    $(\"tab-inbox\").setAttribute(\"aria-selected\", v === \"inbox\"); $(\"tab-sched\").setAttribute(\"aria-selected\", v === \"sched\");\n    if (v === \"sched\") renderSched();\n    window.scrollTo(0, 0);\n  }\n  document.querySelectorAll(\".tab\").forEach(function (b) { b.addEventListener(\"click\", function () { showView(b.dataset.view); }); });\n\n  $(\"refreshBtn\").addEventListener(\"click\", function () { run(\"getState\").then(function (s) { applyState(s); toast(\"새로 불러왔어요\"); }); });\n  setInterval(function () { if (document.visibilityState === \"visible\" && !busy) run(\"getState\").then(applyState, function () {}); }, 60000);\n\n  $(\"inbox\").addEventListener(\"click\", function (e) {\n    var b = e.target.closest(\".sms\"); if (b) openItem(b.dataset.id, true);\n  });\n\n  $(\"pasteForm\").addEventListener(\"submit\", function (e) {\n    e.preventDefault();\n    var txt = $(\"pasteText\").value.trim();\n    if (!txt) { toast(\"내용을 적어 주세요\"); return; }\n    $(\"pasteBtn\").disabled = true; busy = true;\n    run(\"ingestManual\", txt).then(function (res) {\n      $(\"pasteText\").value = \"\"; applyState(res.state); openItem(res.id, true);\n    }).finally(function () { $(\"pasteBtn\").disabled = false; busy = false; });\n  });\n\n  var FIELD = { fDate: \"date\", fTime: \"start\", fDur: \"dur\", fLabel: \"label\", fName: \"customer\", fPhone: \"phone\", fAddr: \"address\" };\n  $(\"result\").addEventListener(\"change\", function (e) {\n    if (!cur || added || !FIELD[e.target.id]) return;\n    var f = FIELD[e.target.id], v = e.target.value;\n    if (f === \"date\") cur.date = v || null;\n    else if (f === \"start\") { if (v) { var p = v.split(\":\"); cur.start = +p[0] * 60 + +p[1]; } else cur.start = null; }\n    else if (f === \"dur\") cur.dur = +v;\n    else { cur[f] = v; $(\"replyText\").value = L.replyText(cur, ev, sel, added); return; }\n    sel = null; reevaluate(false);\n  });\n  $(\"result\").addEventListener(\"input\", function (e) {\n    if (e.target.id === \"replyText\" && $(\"smsLink\")) $(\"smsLink\").href = \"sms:\" + cur.phone + \"?body=\" + encodeURIComponent(e.target.value);\n  });\n  $(\"result\").addEventListener(\"click\", function (e) {\n    if (!cur) return;\n    var evb = e.target.closest(\".ev[data-job]\");\n    if (evb) { var jj = state.jobs.filter(function (x) { return x.id === evb.dataset.job; })[0]; if (jj) toast((jj.fullDay ? \"종일\" : hm(jj.start) + \"–\" + hm(jj.start + jj.dur)) + \" \" + (jj.customer || \"\") + \" · \" + shortLabel(jj.label)); return; }\n    var col = e.target.closest(\".wk-col.pickable\");\n    if (col && !added) {\n      var rc = col.getBoundingClientRect(), m = OPEN + Math.floor((e.clientY - rc.top) / HR * 2) * 30;\n      cur.date = col.dataset.date;\n      if (!cur.full) { cur.start = Math.max(OPEN, Math.min(CLOSE - 30, m)); cur.win = null; }\n      sel = null; reevaluate(false);\n      toast(fmtDate(cur.date) + (cur.full ? \"\" : \" \" + fmtTime(cur.start)) + \"로 다시 확인했어요\");\n      return;\n    }\n    var a = e.target.closest(\".alt\");\n    if (a && !added) { var alt = ev.alts[+a.dataset.alt]; sel = (sel && sel.date === alt.date && sel.start === alt.start) ? null : { date: alt.date, start: alt.start }; renderResult(false); return; }\n    if (e.target.id === \"addBtn\" && sel && !added) {\n      e.target.disabled = true; busy = true;\n      var job = { date: sel.date, start: sel.start, dur: cur.full ? FULL_DUR : cur.dur, fullDay: cur.full, label: cur.label, customer: cur.customer, phone: cur.phone, address: cur.address };\n      run(\"addJob\", job, cur.msgId).then(function (s) {\n        added = true; viewMon = weekMon(sel.date); cur.date = sel.date; if (!cur.full) cur.start = sel.start;\n        var mine = s.jobs.filter(function (x) { return x.date === sel.date && x.start === sel.start; });\n        selJob = mine.length ? mine[mine.length - 1].id : null;\n        applyState(s); toast(\"일정표에 넣었어요\");\n      }, function () { e.target.disabled = false; }).finally(function () { busy = false; });\n      return;\n    }\n    if (e.target.id === \"dismissBtn\") {\n      if (!armedDismiss) { armedDismiss = true; renderResult(false); return; }\n      run(\"setInboxStatus\", cur.msgId, \"dismissed\").then(function (s) { cur = null; applyState(s); toast(\"목록에서 지웠어요\"); });\n      return;\n    }\n    if (e.target.id === \"copyBtn\") {\n      var ta = $(\"replyText\"), txt = ta.value;\n      var fallback = function () { ta.focus(); ta.select(); toast(\"길게 눌러 복사해 주세요\"); };\n      if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(txt).then(function () { toast(\"답장을 복사했어요\"); }, fallback);\n      else fallback();\n    }\n  });\n\n  $(\"settings\").addEventListener(\"click\", function (e) {\n    if (e.target.id === \"copyHook\") {\n      var u = $(\"hookUrl\").textContent;\n      var pick = function () { var r = document.createRange(); r.selectNodeContents($(\"hookUrl\")); var w = window.getSelection(); w.removeAllRanges(); w.addRange(r); toast(\"주소를 길게 눌러 복사해 주세요\"); };\n      if (navigator.clipboard && navigator.clipboard.writeText) navigator.clipboard.writeText(u).then(function () { toast(\"주소를 복사했어요. 전달 앱에 붙여넣으세요\"); }, pick); else pick();\n      return;\n    }\n    if (e.target.id !== \"testKakao\") return;\n    e.target.disabled = true;\n    run(\"sendTestKakao\").then(function (ok) { toast(ok ? \"카카오톡을 확인해 보세요\" : \"보내지 못했어요. 다시 연결해 주세요.\"); }).finally(function () { e.target.disabled = false; });\n  });\n\n  $(\"wkGrid\").addEventListener(\"click\", function (e) {\n    var b = e.target.closest(\".ev[data-job]\"); if (!b) return;\n    selJob = selJob === b.dataset.job ? null : b.dataset.job; armedDel = null; renderSched();\n  });\n  $(\"wkDetail\").addEventListener(\"click\", function (e) {\n    var b = e.target.closest(\".job-del\"); if (!b) return;\n    var id = b.dataset.del;\n    if (armedDel !== id) { armedDel = id; renderSched(); return; }\n    b.disabled = true;\n    run(\"deleteJob\", id).then(function (s) { armedDel = null; selJob = null; applyState(s); toast(\"일정을 지웠어요\"); });\n  });\n  $(\"wkPrev\").addEventListener(\"click\", function () { viewMon = key(addDays(fromKey(viewMon), -7)); selJob = null; renderSched(); });\n  $(\"wkNext\").addEventListener(\"click\", function () { viewMon = key(addDays(fromKey(viewMon), 7)); selJob = null; renderSched(); });\n  $(\"wkToday\").addEventListener(\"click\", function () { viewMon = defaultMon(); selJob = null; renderSched(); });\n\n  var T0 = today0();\n  $(\"todayLabel\").textContent = \"· 오늘 \" + (T0.getMonth() + 1) + \"월 \" + T0.getDate() + \"일 \" + DOW[T0.getDay()] + \"요일\";\n  renderInbox(); renderSched(); renderSettings();\n})();\n</script>\n</body>\n</html>\n";
