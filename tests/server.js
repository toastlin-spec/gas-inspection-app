// 本機測試伺服器：提供 Index.html，並把 google.script.run 轉到模擬後端
const http = require('http');
const fs = require('fs');
const { makeEnv } = require('./gasmock');
const { ctx: G } = makeEnv(__dirname + '/workbook.json');
G.setupSheets_();
const STUB = `<script>
window.google={script:{get run(){let ok=()=>{},fail=()=>{};const h={withSuccessHandler(f){ok=f;return p},withFailureHandler(f){fail=f;return p}};
const p=new Proxy(h,{get(t,k){if(k in t)return t[k];return(...args)=>{fetch('/rpc',{method:'POST',body:JSON.stringify({fn:k,args})}).then(r=>r.json()).then(j=>j.error?fail(new Error(j.error)):ok(j.result))}}});return p}}};
</script>`;
http.createServer((req, res) => {
  if (req.method === 'POST' && req.url === '/rpc') {
    let body = ''; req.on('data', c => body += c); req.on('end', () => {
      const { fn, args } = JSON.parse(body);
      let out;
      try { if (/_$/.test(fn) || typeof G[fn] !== 'function') throw new Error('不允許呼叫 ' + fn); out = { result: G[fn](...args) }; }
      catch (e) { out = { error: e.message }; }
      setTimeout(() => { res.writeHead(200, { 'content-type': 'application/json' }); res.end(JSON.stringify(out)); }, 120);
    });
    return;
  }
  const html = fs.readFileSync(__dirname + '/../src/Index.html', 'utf8').replace('<body>', '<body>' + STUB);
  res.writeHead(200, { 'content-type': 'text/html; charset=utf-8' }); res.end(html);
}).listen(8765, () => console.log('ready'));
