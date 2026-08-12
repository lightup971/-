/* 코어 4종 + 앱 로직을 app_shell.html에 합쳐 단일 실행 파일 생성 */
const fs=require('fs');
const shell=fs.readFileSync('app_shell.html','utf8');
function core(f,globalName){
  let s=fs.readFileSync(f,'utf8');
  s=s.replace(/\n\s*if \(typeof module !== 'undefined' && module\.exports\) module\.exports = API;\n\s*else root\.(\w+) = API;/,
              (_,g)=>`\n  root.${g} = API;`);
  s=s.replace(/\}\)\(typeof self !== 'undefined' \? self : this\);/,'})(window);');
  return s;
}
const cores=['erp_core.js','hometax_core.js','intake_core.js','xlsx_reader.js'].map(f=>core(f)).join('\n');
// hometax_core 내부 zip 노출
const app=fs.readFileSync('app_main.js','utf8');
const m=shell.match(/<title>([\s\S]*?)<\/title>/);
const title=m?m[1]:'수정신고 관리';
const body=shell.replace(/<title>[\s\S]*?<\/title>\s*/,'').replace(/<meta name="viewport"[^>]*>\s*/,'');
const out=`<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${title}</title>
</head>
<body>
${body}
<script>
${cores}
</script>
<script>
${app}
</script>
</body>
</html>
`;
fs.writeFileSync('수정신고_관리앱.html',out,'utf8');
console.log('생성:', '수정신고_관리앱.html', (Buffer.byteLength(out)/1024).toFixed(1)+'KB');
