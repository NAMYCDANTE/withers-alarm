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
