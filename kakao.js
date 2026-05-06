const axios = require('axios');

async function sendKakaoMessage(data) {
  const token = process.env.KAKAO_ACCESS_TOKEN;
  if (!token || token === '여기에_액세스_토큰_입력') {
    console.log('[카카오] 토큰 미설정 - 메시지 전송 건너뜀');
    console.log('[카카오] 전송할 내용:', formatMessage(data));
    return;
  }

  const message = {
    object_type: 'text',
    text: formatMessage(data),
    link: { web_url: 'https://localhost', mobile_web_url: 'https://localhost' }
  };

  try {
    await axios.post(
      'https://kapi.kakao.com/v2/api/talk/memo/default/send',
      `template_object=${encodeURIComponent(JSON.stringify(message))}`,
      {
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/x-www-form-urlencoded'
        }
      }
    );
    console.log('[카카오] 메시지 전송 완료');
  } catch (err) {
    console.error('[카카오] 전송 실패:', err.response?.data || err.message);
  }
}

function formatMessage(data) {
  return `🔔 에어컨 청소 신규 접수

👤 고객명: ${data.customer_name}
📞 연락처: ${data.phone}
📍 주소: ${data.address}
📅 날짜: ${data.date}
⏰ 시간: ${data.time}
🌀 대수: ${data.unit_count}대
📝 메모: ${data.memo || '없음'}

✅ 빠른 확인 부탁드립니다!`;
}

module.exports = { sendKakaoMessage };
