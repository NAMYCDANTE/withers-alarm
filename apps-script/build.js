// 여러 파일을 붙여넣기 쉬운 파일 하나로 합쳐요.
//   node apps-script/build.js
// 결과: apps-script/dist/위더스일정비서.gs  (Apps Script의 Code.gs 에 통째로 붙여넣기)
const fs = require('fs'), path = require('path');
const dir = __dirname, out = path.join(dir, 'dist');
const read = f => fs.readFileSync(path.join(dir, f), 'utf8');
const header = `/**
 * 위더스 일정비서 - 한 파일 버전 (자동 생성 파일, 직접 고치지 말고 apps-script/ 의 원본을 고친 뒤 build.js 실행)
 * 사용법: 구글 시트 > 확장 프로그램 > Apps Script > Code.gs 내용을 모두 지우고 이 파일 전체를 붙여넣기
 */
`;
const gs = header + [read('Code.gs'), read('Logic.gs'), read('Ai.gs')].join('\n\n') +
  '\n\n// ===== 휴대폰 화면 (Index.html) =====\nvar INDEX_HTML = ' + JSON.stringify(read('Index.html')) + ';\n';
fs.mkdirSync(out, { recursive: true });
fs.writeFileSync(path.join(out, '위더스일정비서.gs'), gs);
console.log('wrote dist/위더스일정비서.gs', gs.length, 'bytes');

// 설치 안내 페이지 (코드 복사 버튼에 위 파일을 넣음)
const guide = read('guide.template.html').replace('/*__CODE_JSON__*/""', () => JSON.stringify(gs).replace(/<\//g, '<\\/'));
fs.writeFileSync(path.join(out, 'setup-guide.html'), guide);
console.log('wrote dist/setup-guide.html', guide.length, 'bytes');
