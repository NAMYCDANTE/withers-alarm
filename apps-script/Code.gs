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
