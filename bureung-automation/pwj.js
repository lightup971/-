const { chromium }=require('playwright-core'); const path=require('path');
const F=path.resolve('수정신고_관리앱.html');
(async()=>{
  const b=await chromium.launch({executablePath:'/opt/pw-browsers/chromium-1194/chrome-linux/chrome',args:['--no-sandbox','--allow-file-access-from-files']});
  const ctx=await b.newContext({acceptDownloads:true});
  const p=await ctx.newPage(); const errs=[]; p.on('pageerror',e=>errs.push(e.message));
  await p.goto('file://'+F); await p.waitForTimeout(300);
  console.log('[직접] 배너 hidden =',await p.evaluate(()=>document.getElementById('envwarn').hidden));
  const p2=await ctx.newPage();
  await p2.setContent(`<!doctype html><meta charset=utf-8><iframe src="file://${F}" style="width:1100px;height:800px"></iframe>`);
  await p2.waitForTimeout(1200);
  const fr=p2.frames().find(f=>f!==p2.mainFrame());
  console.log('[iframe] 배너 hidden =',fr?await fr.evaluate(()=>document.getElementById('envwarn').hidden):'프레임없음');
  console.log('JS오류:',errs.length?errs:'없음');
  await b.close();
})().catch(e=>{console.error(e);process.exit(1)});
