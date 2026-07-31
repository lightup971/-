/* erp-generator.html(아티팩트용 조각) → 독립 실행 HTML 문서로 감싸기 */
const fs = require('fs');
const src = fs.readFileSync('erp-generator.html', 'utf8');

// 원본에서 <title>을 뽑아 head로 올림
const m = src.match(/<title>([\s\S]*?)<\/title>/);
const title = m ? m[1] : '수정신고 업로드파일 생성기';
const body = src.replace(/<title>[\s\S]*?<\/title>\s*/, '')
                .replace(/<meta name="viewport"[^>]*>\s*/, '');

const out = `<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${title}</title>
<style>
  /* 최소 리셋 (아티팩트 환경에서 제공되던 것) */
  *,*::before,*::after{box-sizing:border-box}
  body{margin:0}
  button,input,select,textarea{font:inherit;color:inherit}
</style>
</head>
<body>
${body}
</body>
</html>
`;
fs.writeFileSync('수정신고_업로드파일_생성기.html', out, 'utf8');
console.log('생성 완료:', '수정신고_업로드파일_생성기.html', (Buffer.byteLength(out)/1024).toFixed(1)+'KB');
