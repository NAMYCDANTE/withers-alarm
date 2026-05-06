const db = require('./database');

// 날짜+시간 겹치는 접수 감지 (같은 날짜, 1시간 이내)
function checkConflict(date, time) {
  const existing = db.prepare(`
    SELECT * FROM requests
    WHERE date = ? AND status != '취소'
  `).all(date);

  if (existing.length === 0) return [];

  const newHour = parseHour(time);

  return existing.filter(row => {
    const existHour = parseHour(row.time);
    return Math.abs(newHour - existHour) < 1;
  });
}

function parseHour(timeStr) {
  // "오전 10시", "오후 2시", "10:00", "14:00" 등 처리
  if (timeStr.includes(':')) {
    return parseInt(timeStr.split(':')[0]);
  }
  let hour = parseInt(timeStr.replace(/[^0-9]/g, ''));
  if (timeStr.includes('오후') && hour < 12) hour += 12;
  return hour;
}

module.exports = { checkConflict };
