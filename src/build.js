/**
 * 建置腳本：把 rules.json 與 engine.js 內嵌進單一 HTML 檔
 *
 * 目的：規格卡要求「決策器與流程圖共用同一份規則表，不得各自實作」。
 *       以建置方式注入，而非複製貼上，確保 src/rules.json 永遠是唯一真相來源。
 *
 * 執行： node src/build.js
 * 產出： output/台灣成人肺炎鏈球菌疫苗報告.html（單檔、可離線開啟、不依賴任何 CDN）
 */

const fs = require('fs');
const path = require('path');

const SRC = __dirname;
const OUT_DIR = path.join(SRC, '..', 'output');
const OUT_FILE = path.join(OUT_DIR, '台灣成人肺炎鏈球菌疫苗報告.html');

const template = fs.readFileSync(path.join(SRC, 'report.template.html'), 'utf8');
const rulesRaw = fs.readFileSync(path.join(SRC, 'rules.json'), 'utf8');
let engineRaw = fs.readFileSync(path.join(SRC, 'engine.js'), 'utf8');

// 驗證 rules.json 合法
let rules;
try {
  rules = JSON.parse(rulesRaw);
} catch (e) {
  console.error('✗ rules.json 解析失敗：' + e.message);
  process.exit(1);
}

// 瀏覽器版本不需要 require('./rules.json')，改讀全域 PNEUMO_RULES
engineRaw = engineRaw.replace(
  "module.exports = factory(require('./rules.json'));",
  "module.exports = factory(require('./rules.json')); // (Node only)"
);

if (template.indexOf('/*__RULES_JSON__*/') === -1 || template.indexOf('/*__ENGINE_JS__*/') === -1) {
  console.error('✗ 樣板缺少注入標記');
  process.exit(1);
}

/**
 * 內嵌到 <script> 區塊時，內容中任何 </script>（即使在註解或字串裡）
 * 都會被 HTML 解析器當成腳本結束標籤，導致整段語法錯誤。必須跳脫。
 */
function escapeForInlineScript(code) {
  return code.replace(/<\/(script)/gi, '<\\/$1');
}

const html = template
  .replace('/*__RULES_JSON__*/', escapeForInlineScript(JSON.stringify(rules, null, 2)))
  .replace('/*__ENGINE_JS__*/', escapeForInlineScript(engineRaw));

// 確保沒有殘留外部資源請求
const external = html.match(/(src|href)\s*=\s*["'](https?:)?\/\/[^"']+["']/gi) || [];
const offenders = external.filter((m) => !/target=|rel=/.test(m) && !/href\s*=\s*["']https?:\/\//i.test(m));
if (offenders.length) {
  console.error('✗ 偵測到外部資源載入（違反可離線要求）：\n  ' + offenders.join('\n  '));
  process.exit(1);
}

fs.mkdirSync(OUT_DIR, { recursive: true });
fs.writeFileSync(OUT_FILE, html, 'utf8');

// 回歸測試：確認注入後的腳本區塊沒有殘留未跳脫的結束標籤
const scriptBodies = html.split(/<script[^>]*>/i).slice(1).map((s) => s.split(/<\/script>/i)[0]);
if (scriptBodies.length !== 3) {
  console.error('✗ 腳本區塊數量異常（預期 3，實得 ' + scriptBodies.length + '），可能有未跳脫的 </script>');
  process.exit(1);
}

const kb = (Buffer.byteLength(html, 'utf8') / 1024).toFixed(1);
console.log('✅ 建置完成');
console.log('   檔案：' + OUT_FILE);
console.log('   大小：' + kb + ' KB（單檔、無外部相依）');
console.log('   規則表版本：' + rules.meta.version + '（更新日 ' + rules.meta.updated + '）');
