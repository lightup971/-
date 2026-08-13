/* 아티팩트용 HTML 조각 → 독립 실행 문서로 감싸기
   사용: node build_doc.js <입력.html> <출력.html> */
const fs=require('fs');
const [,,src,dst]=process.argv;
if(!src||!dst){ console.error('사용: node build_doc.js <입력> <출력>'); process.exit(1); }
const s=fs.readFileSync(src,'utf8');
const m=s.match(/<title>([\s\S]*?)<\/title>/);
const title=m?m[1].trim():'문서';
const body=s.replace(/<title>[\s\S]*?<\/title>\s*/,'')
            .replace(/<meta name="viewport"[^>]*>\s*/,'');
const out=`<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${title}</title>
<style>
  *,*::before,*::after{box-sizing:border-box}
  body{margin:0}
  @media print{
    body{background:#fff}
    .wrap{max-width:none;padding:0 12mm}
    .cover,.tw,.call,.kpi,.sec{break-inside:avoid}
    h2{break-after:avoid}
  }
</style>
</head>
<body>
${body}
</body>
</html>
`;
fs.writeFileSync(dst,out,'utf8');
console.log(`${dst}  ${(Buffer.byteLength(out)/1024).toFixed(1)}KB`);
