/**
 * 台灣成人肺炎鏈球菌疫苗接種決策引擎
 *
 * 單一真相來源：rules.json
 * 互動決策器與補接種流程圖共用此引擎，不得各自實作判斷邏輯。
 *
 * 用法（瀏覽器）：由 build.js 內嵌後，透過 window.PneumoEngine.decide(input) 呼叫
 * 用法（Node）  ：const { decide } = require('./engine.js')
 */

(function (root, factory) {
  if (typeof module === 'object' && module.exports) {
    module.exports = factory(require('./rules.json'));
  } else {
    root.PneumoEngine = factory(root.PNEUMO_RULES);
  }
})(typeof self !== 'undefined' ? self : this, function (RULES) {
  'use strict';

  /** 依西元年份相減計算年齡，不計月日 */
  function calcAge(birthYear, currentYear) {
    return currentYear - birthYear;
  }

  /** 判定公費資格類別，回傳 category 物件或 null */
  function resolveFundedCategory(ctx) {
    const { age, isIndigenous, fundedRisk } = ctx;
    if (age >= 65) {
      return { id: 'elder65', label: '65 歲（含）以上長者' };
    }
    if (isIndigenous && age >= 55 && age <= 64) {
      return { id: 'indigenous55', label: '55–64 歲原住民' };
    }
    if (age >= 19 && age <= 64 && fundedRisk.length > 0) {
      return { id: 'highrisk19', label: '19–64 歲 IPD 高風險對象' };
    }
    return null;
  }

  /**
   * 是否適用 8 週短間隔
   * 公告逐字：「『IPD高風險對象』、65歲以上『機構住民』及『洗腎患者』接種PCV13/15
   * 間隔至少8週後可接種PCV20或PCV21」——三種身分並列，「65 歲以上」只修飾後兩者，
   * 因此 IPD 高風險對象不分年齡皆適用（65 歲以上的高風險者亦然）。
   */
  function hasShortInterval(ctx) {
    const { age, fundedRisk, isInstitutional, isDialysis } = ctx;
    if (fundedRisk.length > 0) return true;
    if (age >= 65 && (isInstitutional || isDialysis)) return true;
    return false;
  }

  /**
   * 依接種史與身分決定「還要不要打、要等多久」
   * 回傳 { status, requiredMonths, basis, monthsRemaining }
   * status: 'complete' | 'ready' | 'wait' | 'notYetAge65'
   */
  function resolveInterval(ctx) {
    const rule = RULES.intervalRules[ctx.history];
    const { age, monthsSinceLast } = ctx;
    const elapsedAll = monthsSinceLast == null ? 0 : monthsSinceLast;

    /**
     * 對象二（三）：曾完整接種 13 價（或 15 價）＋ 23 價各 1 劑
     * 公告原文限定「先前為 19–64 歲 IPD 高風險對象」才可於滿 65 歲、
     * 間隔 ≥5 年後再接種 1 劑；其餘身分一律視為已完成接種。
     */
    if (rule.boosterRule) {
      if (!ctx.priorHighRisk19to64) {
        return { status: 'complete', requiredMonths: null, basis: rule.completeBasis, boosterNA: true };
      }
      if (age < 65) {
        return { status: 'notYetAge65', requiredMonths: rule.months, basis: rule.basis };
      }
      if (elapsedAll >= rule.months) {
        return { status: 'ready', requiredMonths: rule.months, basis: rule.basis, booster: true };
      }
      return {
        status: 'wait', requiredMonths: rule.months,
        monthsRemaining: rule.months - elapsedAll, basis: rule.basis, booster: true
      };
    }

    if (rule.complete) {
      return { status: 'complete', requiredMonths: null, basis: rule.basis };
    }

    let required = rule.months;
    let usedShort = false;
    if (rule.shortMonths != null && hasShortInterval(ctx)) {
      required = rule.shortMonths;
      usedShort = true;
    }

    if (required === 0) {
      return { status: 'ready', requiredMonths: 0, basis: rule.basis, usedShort };
    }

    // 未曾接種以外的情況都需要 monthsSinceLast
    const elapsed = monthsSinceLast == null ? 0 : monthsSinceLast;
    if (elapsed >= required) {
      return { status: 'ready', requiredMonths: required, basis: rule.basis, usedShort };
    }
    return {
      status: 'wait',
      requiredMonths: required,
      monthsRemaining: required - elapsed,
      basis: rule.basis,
      usedShort
    };
  }

  /** 蒐集警示 */
  function collectWarnings(ctx, fundedCategory, interval) {
    const W = RULES.warnings;
    const out = [];

    const triggers = W.menacwyD.triggerRisks;
    if (ctx.fundedRisk.some((r) => triggers.indexOf(r) !== -1)) {
      out.push({ id: 'menacwyD', level: 'high', text: W.menacwyD.text, note: W.menacwyD.note });
    }
    if (fundedCategory && fundedCategory.id === 'highrisk19') {
      out.push({ id: 'documentation', level: 'info', text: W.documentation.text });
    }
    if (!fundedCategory && ctx.clinicalRisk.length > 0) {
      out.push({ id: 'selfPay', level: 'info', text: W.selfPay.text, source: W.selfPay.source });
    }
    if (interval.booster) {
      out.push({ id: 'boosterEligible', level: 'info', text: W.boosterEligible.text });
    }
    if (interval.boosterNA) {
      out.push({ id: 'boosterNotApplicable', level: 'info', text: W.boosterNotApplicable.text });
    }
    if (interval.usedShort) {
      const reg = RULES.shortIntervalEligibility.registrationCode;
      out.push({ id: 'registrationCode', level: 'info', text: reg.text });
    }
    out.push({ id: 'coadmin', level: 'info', text: W.coadmin.text });
    return out;
  }

  /** 產生 ACIP 對照建議 */
  function resolveAcip(ctx) {
    const isHighRisk = ctx.fundedRisk.length > 0 || ctx.clinicalRisk.length > 0;
    let table = null;
    let tableLabel = '';

    if (ctx.age >= 50) {
      table = RULES.acipComparison.schedule50plus;
      tableLabel = '≥50 歲成人';
    } else if (ctx.age >= 19 && isHighRisk) {
      table = RULES.acipComparison.scheduleHighRisk19to64;
      tableLabel = '高風險 19–64 歲成人';
    } else {
      return {
        applicable: false,
        tableLabel: '—',
        text: 'ACIP 對 <50 歲且無高風險因子之成人未常規建議接種。'
      };
    }
    const entry = table[ctx.history];
    return {
      applicable: true,
      tableLabel,
      optionA: entry.optionA,
      optionB: entry.optionB || null,
      note: entry.note || null
    };
  }

  /**
   * 主判定函式
   * @param {object} input
   * @param {number} input.birthYear      西元出生年
   * @param {number} [input.currentYear]  預設為今年
   * @param {boolean} [input.isIndigenous]
   * @param {string[]} [input.fundedRisk]   公費 5 項適應症 id
   * @param {string[]} [input.clinicalRisk] 臨床高風險（非公費）適應症 id
   * @param {string} input.history          接種史 id
   * @param {number|null} [input.monthsSinceLast] 距上次接種月數
   * @param {boolean} [input.isInstitutional] 機構住民
   * @param {boolean} [input.isDialysis]      洗腎患者
   */
  function decide(input) {
    const currentYear = input.currentYear || new Date().getFullYear();
    const ctx = {
      birthYear: input.birthYear,
      currentYear: currentYear,
      age: calcAge(input.birthYear, currentYear),
      isIndigenous: !!input.isIndigenous,
      fundedRisk: input.fundedRisk || [],
      clinicalRisk: input.clinicalRisk || [],
      history: input.history || 'none',
      monthsSinceLast: input.monthsSinceLast == null ? null : input.monthsSinceLast,
      isInstitutional: !!input.isInstitutional,
      isDialysis: !!input.isDialysis,
      priorHighRisk19to64: !!input.priorHighRisk19to64
    };

    const fundedCategory = resolveFundedCategory(ctx);
    const interval = resolveInterval(ctx);
    const warnings = collectWarnings(ctx, fundedCategory, interval);
    const acip = resolveAcip(ctx);

    // 三段式輸出：該不該打 → 打哪支 → 公費或自費
    let action, headline, detail, vaccine, payment;

    if (interval.status === 'complete') {
      action = 'complete';
      headline = '已完成接種，無需再打';
      detail = interval.basis;
      vaccine = null;
      payment = '—';
    } else if (interval.status === 'notYetAge65') {
      action = 'wait_age';
      headline = '目前無需再打，滿 65 歲後可評估追加';
      detail = '已完整接種 13 價（或 15 價）＋ 23 價各 1 劑。因先前為 19–64 歲 IPD 高風險對象，須滿 65 歲（含）且與前劑間隔滿 5 年後，方可再接種 1 劑 PCV20 或 PCV21。';
      vaccine = null;
      payment = fundedCategory ? '公費（追加時）' : '自費';
    } else if (interval.status === 'wait') {
      action = 'wait_interval';
      headline = `尚未達接種間隔，還需等待約 ${interval.monthsRemaining} 個月`;
      detail = `本情境所需間隔為 ${interval.requiredMonths} 個月${interval.usedShort ? '（適用 8 週短間隔）' : ''}。依據：${interval.basis}`;
      vaccine = RULES.vaccines.offered.slice();
      payment = fundedCategory ? '公費' : '自費';
    } else {
      action = 'vaccinate';
      headline = '現在即可接種 1 劑';
      detail = interval.requiredMonths === 0
        ? '從未接種肺鏈疫苗者無間隔需求。'
        : `已達所需間隔 ${interval.requiredMonths} 個月${interval.usedShort ? '（適用 8 週短間隔）' : ''}。依據：${interval.basis}`;
      vaccine = RULES.vaccines.offered.slice();
      payment = fundedCategory ? '公費' : '自費';
    }

    return {
      input: ctx,
      age: ctx.age,
      eligibility: {
        funded: !!fundedCategory,
        category: fundedCategory ? fundedCategory.id : null,
        categoryLabel: fundedCategory ? fundedCategory.label : '不符現行公費資格',
        selfPayReason: fundedCategory
          ? null
          : ctx.age >= 19 && ctx.age <= 64
            ? '19–64 歲僅限公告列舉的 5 項 IPD 高風險對象具公費資格，其餘（含糖尿病、COPD、慢性肝腎疾病、吸菸等）需自費。'
            : '未達公費年齡條件。'
      },
      recommendation: {
        action: action,
        headline: headline,
        detail: detail,
        vaccine: vaccine,
        vaccineNote: vaccine ? RULES.vaccines.choiceNote : null,
        payment: payment
      },
      acip: acip,
      warnings: warnings,
      basis: {
        tw: RULES.meta.sources.tw,
        acip: RULES.meta.sources.acip
      },
      disclaimer: RULES.disclaimer,
      rulesVersion: RULES.meta.version
    };
  }

  return { decide: decide, calcAge: calcAge, RULES: RULES };
});
