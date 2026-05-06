require('dotenv').config();
const express = require('express');
const bodyParser = require('body-parser');
const db = require('./database');
const { sendKakaoMessage } = require('./kakao');
const { checkConflict } = require('./conflict');

const app = express();
const PORT = process.env.PORT || 3000;

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));
app.use('/pwa', express.static(require('path').join(__dirname, 'pwa')));

// ──────────────────────────────────────────
// 구글 설문지 웹훅 수신
// ──────────────────────────────────────────
app.post('/webhook/form', async (req, res) => {
  try {
    const body = req.body;
    console.log('[웹훅] 수신:', JSON.stringify(body, null, 2));

    // 구글 설문지 응답 파싱
    // Apps Script에서 보내는 형식에 맞게 매핑
    const data = {
      customer_name: body.customer_name || body['고객명'] || '미입력',
      phone:         body.phone        || body['연락처']  || '미입력',
      address:       body.address      || body['주소']    || '미입력',
      date:          body.date         || body['날짜']    || '미입력',
      time:          body.time         || body['시간']    || '미입력',
      unit_count:    parseInt((body.unit_count || body['대수'] || '1').toString().replace(/[^0-9]/g, '')) || 1,
      memo:          body.memo         || body['메모']    || ''
    };

    // 일정 충돌 확인
    const conflicts = checkConflict(data.date, data.time);
    if (conflicts.length > 0) {
      console.log(`[충돌] ${conflicts.length}건 겹치는 일정 있음`);
      data.memo = (data.memo ? data.memo + ' | ' : '') + `⚠️ 일정 충돌 ${conflicts.length}건`;
    }

    // DB 저장
    const stmt = db.prepare(`
      INSERT INTO requests (customer_name, phone, address, date, time, unit_count, memo)
      VALUES (?, ?, ?, ?, ?, ?, ?)
    `);
    const result = stmt.run(
      data.customer_name, data.phone, data.address,
      data.date, data.time, data.unit_count, data.memo
    );
    console.log(`[DB] 저장 완료 (ID: ${result.lastInsertRowid})`);

    // 카카오 알림 전송
    await sendKakaoMessage(data);

    res.json({ success: true, id: result.lastInsertRowid, conflicts: conflicts.length });
  } catch (err) {
    console.error('[오류]', err);
    res.status(500).json({ success: false, error: err.message });
  }
});

// ──────────────────────────────────────────
// 접수 목록 조회
// ──────────────────────────────────────────
app.get('/requests', (req, res) => {
  const rows = db.prepare('SELECT * FROM requests ORDER BY date, time').all();
  res.json(rows);
});

// ──────────────────────────────────────────
// 접수 삭제
// ──────────────────────────────────────────
app.delete('/requests/:id', (req, res) => {
  const { id } = req.params;
  db.prepare('DELETE FROM requests WHERE id = ?').run(id);
  res.json({ success: true });
});

// ──────────────────────────────────────────
// 접수 상태 변경
// ──────────────────────────────────────────
app.patch('/requests/:id/status', (req, res) => {
  const { id } = req.params;
  const { status } = req.body;
  db.prepare('UPDATE requests SET status = ? WHERE id = ?').run(status, id);
  res.json({ success: true });
});

// ──────────────────────────────────────────
// 테스트용: 수동으로 알림 보내기
// ──────────────────────────────────────────
app.post('/test/notify', async (req, res) => {
  const testData = {
    customer_name: '테스트 고객',
    phone: '010-1234-5678',
    address: '서울시 강남구 테스트로 123',
    date: '2026-05-10',
    time: '오전 10시',
    unit_count: 3,
    memo: '테스트 메시지'
  };
  await sendKakaoMessage(testData);
  res.json({ success: true, message: '테스트 알림 전송 완료' });
});

// ──────────────────────────────────────────
// 카카오 OAuth 콜백 (토큰 자동 발급)
// ──────────────────────────────────────────
app.get('/kakao-callback', async (req, res) => {
  const { code, error } = req.query;
  if (error) {
    return res.send(`<h2>오류: ${error}</h2>`);
  }
  if (!code) {
    return res.send('<h2>코드가 없습니다.</h2>');
  }

  const axios = require('axios');
  const fs = require('fs');
  const path = require('path');

  try {
    const response = await axios.post(
      'https://kauth.kakao.com/oauth/token',
      new URLSearchParams({
        grant_type: 'authorization_code',
        client_id: 'afbc5508d9a2a9377f6a60bade33c963',
        redirect_uri: `https://surfer-relation-makeshift.ngrok-free.dev/kakao-callback`,
        code: code
      }).toString(),
      { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
    );

    const { access_token, refresh_token } = response.data;

    // .env 파일에 토큰 저장
    const envPath = path.join(__dirname, '.env');
    let envContent = fs.readFileSync(envPath, 'utf8');
    envContent = envContent.replace(
      /KAKAO_ACCESS_TOKEN=.*/,
      `KAKAO_ACCESS_TOKEN=${access_token}`
    );
    fs.writeFileSync(envPath, envContent);

    console.log('[카카오] 토큰 발급 완료!');
    console.log('Access Token:', access_token);

    res.send(`
      <h2>✅ 카카오 토큰 발급 완료!</h2>
      <p>서버를 재시작하면 카카오 알림이 활성화됩니다.</p>
      <p><b>서버를 Ctrl+C로 종료하고 npm start로 다시 시작해주세요.</b></p>
    `);
  } catch (err) {
    console.error('[카카오] 토큰 발급 실패:', err.response?.data || err.message);
    res.send(`<h2>오류 발생</h2><pre>${JSON.stringify(err.response?.data, null, 2)}</pre>`);
  }
});

app.listen(PORT, () => {
  console.log(`\n✅ 위더스환경 알림 서버 시작`);
  console.log(`📡 포트: ${PORT}`);
  console.log(`🔗 웹훅 URL: http://localhost:${PORT}/webhook/form`);
  console.log(`📋 접수 목록: http://localhost:${PORT}/requests\n`);
});
