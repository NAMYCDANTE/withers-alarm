// ============================================================
// 위더스환경 에어컨 청소 - 구글 설문지 자동 알림
// ============================================================
// 사용법:
// 1. 구글 시트 열기 → 확장 프로그램 → Apps Script
// 2. 이 코드 전체 붙여넣기
// 3. WEBHOOK_URL을 ngrok 주소로 변경
// 4. 저장 후 트리거 설정 (아래 설명 참고)
// ============================================================

// ▼ 여기에 ngrok 주소 입력 (예: https://abc123.ngrok.io)
var WEBHOOK_URL = "https://여기에_ngrok_주소_입력/webhook/form";

// ============================================================
// 설문지 항목 이름 (구글 설문지의 질문 제목과 정확히 일치시켜야 함)
// ============================================================
var FIELD_MAP = {
  customer_name: "고객명",       // 설문지 질문 제목
  phone:         "연락처",
  address:       "주소",
  date:          "희망 날짜",
  time:          "희망 시간",
  unit_count:    "에어컨 대수",
  memo:          "요청사항"
};

// ============================================================
// 설문지 제출 시 자동 실행되는 함수
// ============================================================
function onFormSubmit(e) {
  try {
    var responses = e.namedValues;

    // 설문지 응답 파싱
    var data = {
      customer_name: getValue(responses, FIELD_MAP.customer_name),
      phone:         getValue(responses, FIELD_MAP.phone),
      address:       getValue(responses, FIELD_MAP.address),
      date:          getValue(responses, FIELD_MAP.date),
      time:          getValue(responses, FIELD_MAP.time),
      unit_count:    getValue(responses, FIELD_MAP.unit_count) || "1",
      memo:          getValue(responses, FIELD_MAP.memo)
    };

    Logger.log("전송 데이터: " + JSON.stringify(data));

    // 웹훅 서버로 전송
    var options = {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(data),
      muteHttpExceptions: true
    };

    var response = UrlFetchApp.fetch(WEBHOOK_URL, options);
    var result = JSON.parse(response.getContentText());

    Logger.log("서버 응답: " + JSON.stringify(result));

    if (result.conflicts > 0) {
      // 일정 충돌 시 이메일 추가 알림 (선택사항)
      Logger.log("⚠️ 일정 충돌 " + result.conflicts + "건 발생!");
    }

  } catch (err) {
    Logger.log("오류 발생: " + err.message);
    // 오류 발생 시 이메일로 알림
    MailApp.sendEmail(
      Session.getActiveUser().getEmail(),
      "[위더스환경] 알림 전송 오류",
      "오류 내용: " + err.message
    );
  }
}

// 응답값 추출 헬퍼 함수
function getValue(responses, fieldName) {
  if (!responses[fieldName]) return "";
  var val = responses[fieldName];
  return Array.isArray(val) ? val[0] : val;
}

// ============================================================
// 트리거 자동 등록 함수 (최초 1회만 실행)
// ============================================================
function setupTrigger() {
  // 기존 트리거 삭제
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "onFormSubmit") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  // 새 트리거 등록 (설문지 제출 시 실행)
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ScriptApp.newTrigger("onFormSubmit")
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();

  Logger.log("✅ 트리거 등록 완료!");
}

// ============================================================
// 테스트 함수 (실제 설문지 없이 테스트)
// ============================================================
function testWebhook() {
  var fakeEvent = {
    namedValues: {
      "고객명":      ["홍길동"],
      "연락처":      ["010-1234-5678"],
      "주소":        ["서울시 강남구 테헤란로 123"],
      "희망 날짜":   ["2026-05-15"],
      "희망 시간":   ["오전 10시"],
      "에어컨 대수": ["2"],
      "요청사항":    ["벽걸이형 2대"]
    }
  };
  onFormSubmit(fakeEvent);
}
