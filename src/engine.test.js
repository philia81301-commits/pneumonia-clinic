/**
 * 規則引擎驗證
 *
 * 對應規格卡驗收條件：
 *   「互動決策器可跑通全部輸入組合、無『無建議』死角，且抽測 5 種情境與流程圖答案一致」
 *
 * 執行： node src/engine.test.js
 */

const { decide } = require('./engine.js');

const YEAR = 2026;
let failures = 0;
let checks = 0;

function assert(cond, label, extra) {
  checks++;
  if (!cond) {
    failures++;
    console.error(`  ✗ ${label}${extra ? '\n      ' + extra : ''}`);
  }
}

// ─────────────────────────────────────────────────────────────
// 1. 窮舉：所有輸入組合都必須產出完整建議，不得有死角
// ─────────────────────────────────────────────────────────────
console.log('【1】窮舉全部輸入組合，檢查無「無建議」死角');

const histories = ['none', 'ppv23', 'pcv13', 'pcv15', 'pcv13_ppv23', 'pcv15_ppv23', 'pcv20', 'pcv21'];
const monthsOptions = [null, 0, 2, 6, 12, 60, 84];
const validActions = ['vaccinate', 'wait_interval', 'wait_age', 'complete'];

let combos = 0;
for (let age = 18; age <= 95; age++) {
  const birthYear = YEAR - age;
  for (const isIndigenous of [false, true]) {
    for (const fundedRisk of [[], ['asplenia'], ['cochlearImplant'], ['immunodeficiency', 'transplant']]) {
      for (const clinicalRisk of [[], ['diabetes'], ['lungDisease', 'smoking']]) {
        for (const history of histories) {
          for (const monthsSinceLast of monthsOptions) {
            for (const isInstitutional of [false, true]) {
              for (const priorHighRisk19to64 of [false, true]) {
                const isDialysis = priorHighRisk19to64;   // 兩個布林維度交錯覆蓋，避免組合數爆炸
                combos++;
                let r;
                try {
                  r = decide({
                    birthYear, currentYear: YEAR, isIndigenous, fundedRisk, clinicalRisk,
                    history, monthsSinceLast, isInstitutional, isDialysis, priorHighRisk19to64
                  });
                } catch (e) {
                  failures++;
                  console.error(`  ✗ 拋出例外 age=${age} history=${history}: ${e.message}`);
                  continue;
                }
                const tag = `age=${age} hist=${history} m=${monthsSinceLast} funded=[${fundedRisk}] indig=${isIndigenous}`;
                if (validActions.indexOf(r.recommendation.action) === -1) {
                  failures++; console.error(`  ✗ 未知 action：${r.recommendation.action}｜${tag}`);
                }
                if (!r.recommendation.headline) { failures++; console.error(`  ✗ headline 空白｜${tag}`); }
                if (!r.recommendation.detail)   { failures++; console.error(`  ✗ detail 空白｜${tag}`); }
                if (!r.recommendation.payment)  { failures++; console.error(`  ✗ payment 空白｜${tag}`); }
                if (r.age !== age)              { failures++; console.error(`  ✗ 年齡計算錯誤｜${tag}`); }
                if (!r.eligibility.categoryLabel) { failures++; console.error(`  ✗ 資格標籤空白｜${tag}`); }
                if (r.acip.applicable && !r.acip.optionA) { failures++; console.error(`  ✗ ACIP 建議空白｜${tag}`); }
                if (!Array.isArray(r.warnings)) { failures++; console.error(`  ✗ warnings 非陣列｜${tag}`); }
                if (!r.disclaimer)              { failures++; console.error(`  ✗ 免責聲明遺失｜${tag}`); }
              }
            }
          }
        }
      }
    }
  }
}
checks += combos;
console.log(`  共測試 ${combos.toLocaleString()} 種組合\n`);

// ─────────────────────────────────────────────────────────────
// 2. 抽測門診情境：必須與補接種流程圖答案一致
// ─────────────────────────────────────────────────────────────
console.log('【2】抽測門診情境，比對流程圖預期答案');

const scenarios = [
  {
    name: '① 70 歲、從未接種、無特殊病史',
    input: { birthYear: 1956, history: 'none' },
    expect: { funded: true, action: 'vaccinate', payment: '公費', category: 'elder65' }
  },
  {
    name: '② 68 歲、只打過 PPV23、距今 6 個月',
    input: { birthYear: 1958, history: 'ppv23', monthsSinceLast: 6 },
    expect: { funded: true, action: 'wait_interval', payment: '公費', monthsRemaining: 6 }
  },
  {
    name: '③ 62 歲糖尿病、從未接種（台灣不符公費）',
    input: { birthYear: 1964, history: 'none', clinicalRisk: ['diabetes'] },
    expect: { funded: false, action: 'vaccinate', payment: '自費', acipApplicable: true }
  },
  {
    name: '④ 45 歲脾臟功能缺損、打過 PCV13 距今 3 個月（適用 8 週短間隔）',
    input: { birthYear: 1981, history: 'pcv13', monthsSinceLast: 3, fundedRisk: ['asplenia'] },
    expect: { funded: true, action: 'vaccinate', payment: '公費', category: 'highrisk19', hasWarning: 'menacwyD' }
  },
  {
    name: '⑤ 67 歲、舊制 PCV13＋PPV23、距前劑 7 年、非高風險出身（視為已完成）',
    input: { birthYear: 1959, history: 'pcv13_ppv23', monthsSinceLast: 84 },
    expect: { funded: true, action: 'complete', payment: '—', hasWarning: 'boosterNotApplicable' }
  },
  {
    name: '⑤b 同上，但先前為 19–64 歲 IPD 高風險對象（對象二（三）可追加）',
    input: { birthYear: 1959, history: 'pcv13_ppv23', monthsSinceLast: 84, priorHighRisk19to64: true },
    expect: { funded: true, action: 'vaccinate', payment: '公費', hasWarning: 'boosterEligible' }
  },
  {
    name: '⑤c 同上但距前劑僅 3 年（未滿 5 年）',
    input: { birthYear: 1959, history: 'pcv13_ppv23', monthsSinceLast: 36, priorHighRisk19to64: true },
    expect: { funded: true, action: 'wait_interval', monthsRemaining: 24 }
  },
  {
    name: '⑤d PCV15＋PPV23 亦適用同一條（公告寫「13價(或15價)」）',
    input: { birthYear: 1959, history: 'pcv15_ppv23', monthsSinceLast: 84, priorHighRisk19to64: true },
    expect: { funded: true, action: 'vaccinate', payment: '公費' }
  },
  {
    name: '⑥ 70 歲機構住民、打過 PCV13 距今 3 個月（8 週短間隔）',
    input: { birthYear: 1956, history: 'pcv13', monthsSinceLast: 3, isInstitutional: true },
    expect: { funded: true, action: 'vaccinate', payment: '公費' }
  },
  {
    name: '⑦ 70 歲一般長者、打過 PCV13 距今 3 個月（不適用短間隔，須滿 1 年）',
    input: { birthYear: 1956, history: 'pcv13', monthsSinceLast: 3 },
    expect: { funded: true, action: 'wait_interval', payment: '公費', monthsRemaining: 9 }
  },
  {
    name: '⑧ 58 歲原住民、從未接種',
    input: { birthYear: 1968, history: 'none', isIndigenous: true },
    expect: { funded: true, action: 'vaccinate', payment: '公費', category: 'indigenous55' }
  },
  {
    name: '⑨ 66 歲、已打過 PCV21',
    input: { birthYear: 1960, history: 'pcv21' },
    expect: { funded: true, action: 'complete', payment: '—' }
  },
  {
    name: '⑩ 55 歲 IPD 高風險、舊制兩劑（未滿 65 歲，尚不可追加）',
    input: { birthYear: 1971, history: 'pcv13_ppv23', monthsSinceLast: 84, fundedRisk: ['csfLeak'], priorHighRisk19to64: true },
    expect: { funded: true, action: 'wait_age' }
  }
];

for (const s of scenarios) {
  const r = decide(Object.assign({ currentYear: YEAR }, s.input));
  const e = s.expect;
  console.log(`\n  ${s.name}`);
  console.log(`    → ${r.recommendation.headline}｜${r.eligibility.categoryLabel}｜${r.recommendation.payment}`);

  if (e.funded !== undefined) assert(r.eligibility.funded === e.funded, `${s.name}：公費資格應為 ${e.funded}`, `實得 ${r.eligibility.funded}`);
  if (e.action) assert(r.recommendation.action === e.action, `${s.name}：action 應為 ${e.action}`, `實得 ${r.recommendation.action}`);
  if (e.payment) assert(r.recommendation.payment === e.payment, `${s.name}：payment 應為 ${e.payment}`, `實得 ${r.recommendation.payment}`);
  if (e.category) assert(r.eligibility.category === e.category, `${s.name}：category 應為 ${e.category}`, `實得 ${r.eligibility.category}`);
  if (e.monthsRemaining !== undefined) {
    const m = /還需等待約 (\d+) 個月/.exec(r.recommendation.headline);
    assert(m && Number(m[1]) === e.monthsRemaining, `${s.name}：剩餘月數應為 ${e.monthsRemaining}`, `實得 ${m ? m[1] : '無'}`);
  }
  if (e.hasWarning) assert(r.warnings.some((w) => w.id === e.hasWarning), `${s.name}：應含警示 ${e.hasWarning}`);
  if (e.acipApplicable !== undefined) assert(r.acip.applicable === e.acipApplicable, `${s.name}：ACIP 適用性應為 ${e.acipApplicable}`);
}

// ─────────────────────────────────────────────────────────────
// 3. 台美落差必須被正確呈現
// ─────────────────────────────────────────────────────────────
console.log('\n\n【3】台美落差檢查');

const gap = decide({ birthYear: 1970, currentYear: YEAR, history: 'none', clinicalRisk: ['diabetes'] }); // 56 歲糖尿病
console.log(`  56 歲糖尿病：台灣 ${gap.recommendation.payment}／ACIP ${gap.acip.applicable ? '建議接種：' + gap.acip.optionA : '未建議'}`);
assert(gap.eligibility.funded === false, '56 歲糖尿病在台灣不應具公費資格');
assert(gap.acip.applicable === true, '56 歲在 ACIP 應適用（≥50 歲）');

const gap2 = decide({ birthYear: 1961, currentYear: YEAR, history: 'none', clinicalRisk: ['diabetes'] }); // 65 歲
console.log(`  65 歲同一位病人：台灣 ${gap2.recommendation.payment}`);
assert(gap2.eligibility.funded === true, '65 歲應轉為公費');

// ─────────────────────────────────────────────────────────────
console.log('\n' + '─'.repeat(60));
if (failures === 0) {
  console.log(`✅ 全部通過（${checks.toLocaleString()} 項檢查，0 失敗）`);
  process.exit(0);
} else {
  console.log(`❌ ${failures} 項失敗`);
  process.exit(1);
}
