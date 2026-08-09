/**
 * 台灣成人肺炎鏈球菌疫苗簡報產生器
 *
 * 原則（規格卡）：一頁一訊息、張數不設上限、避免單頁擁擠
 * 版面預設（全域）：淺色底、內文 20–24pt 粗體、留白約 30%、文案精簡不斷行
 * 數字一律取自 data/01–07，與 HTML 報告保持一致
 *
 * 執行： node src/deck.js
 */

const pptxgen = require('pptxgenjs');
const path = require('path');
const fs = require('fs');
const RULES = require('./rules.json');   // 政策文字取自單一規則表，避免簡報與決策器各自寫死

const pres = new pptxgen();
pres.layout = 'LAYOUT_WIDE';          // 13.3 × 7.5 英吋
pres.author = '潘醫師';
pres.title = '台灣成人肺炎鏈球菌疫苗：血清型、證據與政策';

const W = 13.3, H = 7.5;
const M = 0.7;                         // 邊界留白
const CW = W - M * 2;                  // 內容寬

// ── 配色（深青主色，暖橘警示；不用一般企業藍）──
const C = {
  deep:  '0F4C5C',   // 主色：深青
  mid:   '1F6F8B',   // 次色
  soft:  'E4EFF3',   // 淺青底
  bg:    'FFFFFF',
  ink:   '15242C',
  ink2:  '4A5B65',
  warn:  'B45309',   // 警示（暖橘）
  warnBg:'FDF3E3',
  bad:   '9E2A2B',   // 紅線
  badBg: 'FBECEC',
  ok:    '1E7A52',
  okBg:  'E7F4EE',
  line:  'D3E0E6'
};

/* ── 疫苗語意色：世代由舊到新、色階由淺到深 ──
   ★ 約束：這五個色只代表疫苗，且只用在**資料呈現**（圖表、矩陣、表格的疫苗欄）。
     投影片本身的裝飾（章節頁編號、標題）不受此約束，但資料區塊內不得把這幾色挪作他用，
     也不得在圖表裡臨時另配一套顏色。
   ★ HTML 報告 report.template.html 的 --ppv23/--pcv13/--pcv20/--pcv21/--shared 必須與此同值。 */
const V = {
  PPV23:  'A9C9D4',
  PCV13:  '6BA3B5',
  PCV15:  '6BA3B5',   // 與 PCV13 同世代，共用色階
  PCV20:  '2E7F9B',
  PCV21:  '0F4C5C',
  shared: '5F7783'    // 「兩者共有」用中性色，不歸屬任何一支（夠深，可直接當文字色）
};

const F = 'Microsoft JhengHei';

let n = 0;
const notes = [];

/* ─────────── 版面元件 ─────────── */

/* 縮寫展開：投影片常被單張抽出來用，聽眾也可能中途進場，
   所以「只在首次出現展開」不夠 —— 有出現該縮寫的每一張都要在頁尾標。
   下面用自動偵測，不靠人記得加；要新增縮寫只要往 ABBR 加一行。 */
const ABBR = {
  IPD: 'IPD ＝ 侵襲性肺炎鏈球菌感染症'
};

function slide(opts) {
  opts = opts || {};
  const s = pres.addSlide();
  s.background = { color: opts.dark ? C.deep : C.bg };
  // 累積本張的所有文字，供 foot() 判斷要標哪些縮寫
  s.__txt = '';
  const _addText = s.addText.bind(s);
  s.addText = function (t, o) {
    s.__txt += typeof t === 'string' ? t : JSON.stringify(t);
    return _addText(t, o);
  };
  const _addTable = s.addTable.bind(s);
  s.addTable = function (rows, o) {
    s.__txt += JSON.stringify(rows);
    return _addTable(rows, o);
  };
  n++;
  return s;
}

/** 章節標記 + 標題（淺色頁） */
function head(s, title, kicker) {
  if (kicker) {
    s.addText(kicker, {
      x: M, y: 0.42, w: CW, h: 0.32, fontFace: F, fontSize: 14, bold: true,
      color: C.mid, charSpacing: 2
    });
  }
  s.addText(title, {
    x: M, y: kicker ? 0.78 : 0.6, w: CW, h: 1.0, fontFace: F, fontSize: 36, bold: true,
    color: C.ink, valign: 'top'
  });
}

/** 頁碼與出處 */
function foot(s, src) {
  // 縮寫併進出處行。不另起一行是因為版面下緣常被 callout 佔住，
  // 另起一行會撞到內容（實測 S3／S37／S38／S45 都會）。
  const abbr = Object.keys(ABBR)
    .filter(k => (s.__txt || '').indexOf(k) >= 0)
    .map(k => ABBR[k]).join('　');
  const line = [abbr, src].filter(Boolean).join('　｜　');
  if (line) {
    s.addText(line, {
      x: M, y: H - 0.62, w: CW - 0.8, h: 0.34, fontFace: F, fontSize: 11,
      color: C.ink2, valign: 'middle'
    });
  }
  s.addText(String(n), {
    x: W - M - 0.6, y: H - 0.62, w: 0.6, h: 0.34, fontFace: F, fontSize: 11,
    color: C.ink2, align: 'right', valign: 'middle'
  });
}

/** 卡片（無邊條，用淡底＋圓角，符合避免裝飾色條的要求） */
function card(s, x, y, w, h, fill) {
  s.addShape(pres.ShapeType.roundRect, {
    x: x, y: y, w: w, h: h, rectRadius: 0.12,
    fill: { color: fill || C.soft }, line: { color: fill ? fill : C.soft, width: 0 }
  });
}

/** 圓形編號 */
function numCircle(s, x, y, d, label, bg, fg) {
  s.addShape(pres.ShapeType.ellipse, {
    x: x, y: y, w: d, h: d, fill: { color: bg || C.mid }, line: { width: 0 }
  });
  s.addText(label, {
    x: x, y: y, w: d, h: d, fontFace: F, fontSize: 18, bold: true,
    color: fg || 'FFFFFF', align: 'center', valign: 'middle', margin: 0
  });
}

/** 大數字區塊 */
function bigStat(s, x, y, w, value, label, color) {
  s.addText(value, {
    x: x, y: y, w: w, h: 1.25, fontFace: F, fontSize: 60, bold: true,
    color: color || C.deep, align: 'center', valign: 'middle', margin: 0
  });
  s.addText(label, {
    x: x, y: y + 1.25, w: w, h: 0.72, fontFace: F, fontSize: 16, bold: true,
    color: C.ink2, align: 'center', valign: 'top', margin: 0
  });
}

/** 條列（內文 20–24pt 粗體） */
function bullets(s, x, y, w, items, size) {
  const arr = items.map((t, i) => ({
    text: t,
    options: { bullet: true, breakLine: i !== items.length - 1 }
  }));
  s.addText(arr, {
    x: x, y: y, w: w, h: 0.1 + items.length * 0.62, fontFace: F,
    fontSize: size || 21, bold: true, color: C.ink, paraSpaceAfter: 10, valign: 'top'
  });
}

/** 提示框 */
function callout(s, x, y, w, h, text, kind, size) {
  const map = { warn: [C.warnBg, C.warn], bad: [C.badBg, C.bad], ok: [C.okBg, C.ok], info: [C.soft, C.deep] };
  const pair = map[kind || 'info'];
  card(s, x, y, w, h, pair[0]);
  s.addText(text, {
    x: x + 0.3, y: y + 0.18, w: w - 0.6, h: h - 0.36, fontFace: F, fontSize: size || 20, bold: true,
    color: pair[1], valign: 'middle'
  });
}

/**
 * 表格
 * totalH：表格要佔滿的總高度（英吋）。給定時平均分配列高，
 *         讓有欄框的版面確實佔到約 70% 版面，而不是縮在中間。
 */
function table(s, x, y, w, rows, colW, fontSize, totalH) {
  const opt = {
    x: x, y: y, w: w, colW: colW,
    fontFace: F, fontSize: fontSize || 16, color: C.ink,
    border: { type: 'solid', color: C.line, pt: 1 },
    fill: { color: 'FFFFFF' },
    valign: 'middle', autoPage: false
  };
  // totalH 可給數字（平均分配）或陣列（逐列指定，用於某列文字特別長時）
  if (Array.isArray(totalH)) opt.rowH = totalH;
  else if (totalH) opt.rowH = totalH / rows.length;
  s.addTable(rows, opt);
}

/** 內容區的標準底線：表格／卡片填滿到這裡，才達到約 70% 版面 */
const CONTENT_BOTTOM = 6.55;
function th(t) { return { text: t, options: { bold: true, fill: { color: C.soft }, color: C.deep } }; }
function td(t, opt) { return { text: t, options: Object.assign({ bold: false }, opt || {}) }; }

function note(s, txt) { s.addNotes(txt); notes.push(txt); }

/* ═══════════════════════════════════════════════════════
   封面
   ═══════════════════════════════════════════════════════ */
{
  const s = slide({ dark: true });
  s.addText('台灣成人肺炎鏈球菌疫苗', {
    x: M, y: 2.0, w: CW, h: 0.9, fontFace: F, fontSize: 44, bold: true, color: 'FFFFFF'
  });
  s.addText('血清型涵蓋、效力證據等級與公費政策', {
    x: M, y: 3.0, w: CW, h: 0.7, fontFace: F, fontSize: 26, bold: true, color: 'CFE3EA'
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 4.1, w: 6.4, h: 0.72, rectRadius: 0.12,
    fill: { color: '17607A' }, line: { width: 0 }
  });
  s.addText('PPV23 ｜ PCV13 ｜ PCV20 ｜ PCV21', {
    x: M, y: 4.1, w: 6.4, h: 0.72, fontFace: F, fontSize: 20, bold: true,
    color: 'FFFFFF', align: 'center', valign: 'middle', margin: 0
  });
  s.addText('高雄榮民總醫院　家庭醫學部　潘湘如 醫師', {
    x: M, y: 5.15, w: CW, h: 0.5, fontFace: F, fontSize: 22, bold: true, color: 'FFFFFF'
  });
  s.addText('資料查證日 2026-08-07　｜　本簡報不使用任何病人資料', {
    x: M, y: 6.3, w: CW, h: 0.4, fontFace: F, fontSize: 14, color: '9DC2CE'
  });
  note(s, '本簡報所有數字均可回溯至原始文獻或官方公告，出處列於各頁下方與最末頁。');
}

/* ═══════════ 一、政策速報 ═══════════ */
{
  const s = slide();
  head(s, '新制 2026 年 8 月 10 日上路', '政策速報');
  bigStat(s, M, 2.2, 3.9, '1 劑', '搞定\n取代舊制兩劑');
  s.addShape(pres.ShapeType.rightArrow, {
    x: 4.8, y: 2.65, w: 1.0, h: 0.5, fill: { color: C.line }, line: { width: 0 }
  });
  card(s, 6.0, 2.2, CW - 5.3, 2.0);
  s.addText([
    { text: '舊制　PCV13 ＋ PPV23（兩劑）', options: { breakLine: true, color: C.ink2 } },
    { text: '新制　PCV20 或 PCV21（二擇一，一劑）', options: { color: C.deep } }
  ], {
    x: 6.3, y: 2.45, w: CW - 5.9, h: 1.5, fontFace: F, fontSize: 22, bold: true, valign: 'middle'
  });
  callout(s, M, 4.6, CW, 0.95, 'PPV23 於新制實施後不再提供公費接種，疾管署辦理回收作業', 'bad');
  foot(s, '疾管署致醫界通函第 612 號（2026-07-28）｜查詢日 2026-08-07');
  note(s, '這是全場最需要先講清楚的一件事：門診從 8/10 起不再是兩劑流程。');
}

{
  const s = slide();
  head(s, '公費對象：三類', '政策');
  const rows = [
    [th('類別'), th('條件')],
    [td('長者'), td('65 歲（含）以上', { bold: true, color: C.deep })],
    [td('原住民'), td('55–64 歲', { bold: true, color: C.deep })],
    // 全場第一次出現 IPD，在此展開全名（之後每一張有 IPD 的投影片由 foot() 自動在頁尾標）
    [td('IPD 高風險\n（侵襲性肺炎鏈球菌感染症）'), td('19–64 歲，且符合公告列舉的 5 項之一：\n① 脾臟功能缺損　② 先天或後天免疫功能不全　③ 人工耳植入\n④ 腦脊髓液滲漏　⑤「一年內」接受免疫抑制劑或放射治療的惡性腫瘤者及器官移植者', { bold: true, color: C.deep })]
  ];
  table(s, M, 2.1, CW, rows, [2.9, CW - 2.9], 18, [0.62, 0.62, 0.62, 1.62]);
  callout(s, M, 5.75, CW, 1.2,
    '「IPD 高風險對象」僅限上列 5 項。\n糖尿病、COPD、慢性肝腎疾病、吸菸等不在其中，65 歲前需自費。', 'bad', 21);
  foot(s, '成人肺炎鏈球菌疫苗接種須知（自 115 年 8 月 10 日起適用）｜查詢日 2026-08-08');
  note(s, '這裡就要把 5 項講明白。「高風險」三個字很容易讓人自動聯想到糖尿病與 COPD，那是錯的，而且會一路錯到門診。第 5 項是把惡性腫瘤與器官移植併為同一項，且限一年內接受治療者。');
}

/* ═══════════ 二、台灣流病 ═══════════ */
{
  const s = slide({ dark: true });
  s.addText('01', { x: M, y: 2.3, w: 2.0, h: 1.0, fontFace: F, fontSize: 54, bold: true, color: '4E93A8' });
  s.addText('台灣流行病學現況', {
    x: M, y: 3.3, w: CW, h: 0.9, fontFace: F, fontSize: 40, bold: true, color: 'FFFFFF'
  });
  s.addText('為什麼疾管署要在此時改制', {
    x: M, y: 4.2, w: CW, h: 0.6, fontFace: F, fontSize: 22, bold: true, color: 'A8CBD7'
  });
  note(s, '接下來三頁各講一個數字，不要一次全丟。');
}

{
  const s = slide();
  head(s, 'IPD 發生率在疫情後回升', '台灣流病');
  bigStat(s, M, 2.3, 5.4, '1.25', '2024 年每十萬人口\n侵襲性肺炎鏈球菌感染症發生率', C.deep);
  card(s, 6.6, 2.2, CW - 5.9, 2.7);
  s.addText('公費疫苗施打後\n曾明顯下降，\n新冠疫情解封後\n再度回升。', {
    x: 6.9, y: 2.35, w: CW - 6.5, h: 2.4, fontFace: F, fontSize: 28, bold: true,
    color: C.ink, valign: 'middle'
  });
  foot(s, '內科學誌 2025;36(6):381-386｜查詢日 2026-08-07');
  note(s, '一頁一數字，讓聽眾記得住 1.25。');
}

{
  const s = slide();
  head(s, '病例集中在中高齡', '台灣流病');
  bigStat(s, M, 2.3, 5.4, '約 7 成', '病例發生於\n50 歲以上成人', C.deep);
  card(s, 6.6, 2.2, CW - 5.9, 2.7, C.warnBg);
  s.addText('50 歲已經是\n風險轉折點，\n但台灣公費\n從 65 歲才開始。', {
    x: 6.9, y: 2.35, w: CW - 6.5, h: 2.4, fontFace: F, fontSize: 28, bold: true,
    color: C.warn, valign: 'middle'
  });
  foot(s, '內科學誌 2025;36(6):381-386｜查詢日 2026-08-07');
  note(s, '這頁埋一個伏筆：台美起始年齡差 15 歲，後面會回來講。');
}

{
  const s = slide();
  head(s, '真正的問題是覆蓋率', '台灣流病');
  bigStat(s, M, 2.3, 5.4, '約 12%', '2024 年完整接種兩劑\n（PCV13/15 ＋ PPV23）比率', C.bad);
  card(s, 6.6, 2.2, CW - 5.9, 2.7, C.warnBg);
  s.addText('舊制兩劑\n完成率低，\n正是改為\n「1 劑搞定」的動機。', {
    x: 6.9, y: 2.35, w: CW - 6.5, h: 2.4, fontFace: F, fontSize: 28, bold: true,
    color: C.warn, valign: 'middle'
  });
  callout(s, M, 4.95, CW, 0.8, '疾管署另公布 31%→36%：母體是 65 歲以上長者的「接種完成率」，與本頁定義不同', 'warn');
  foot(s, '內科學誌 2025;36(6):381-386｜疾管署新聞稿 2026-07-28｜查詢日 2026-08-07、2026-08-09');
  note(s, '被問到 36% 時要能立刻分辨：本頁 12% 是 2024 年全成人、舊制兩劑都打完；'
       + '疾管署的 31%→36% 是 65 歲以上長者的接種完成率，而且新制打 1 劑就算完成。'
       + '不要講成「從 12% 提升到 36%」，那是三個不同定義串成的假趨勢。');
}

{
  const s = slide();
  head(s, 'IPD 病例逐年上升，死亡數未同步增加', '台灣流病');
  const YR = ['2021', '2022', '2023', '2024', '2025'];
  s.addChart([
    {
      type: pres.ChartType.bar,
      data: [{ name: '確定病例數', labels: YR, values: [196, 200, 287, 315, 347] }],
      options: {
        barDir: 'col', chartColors: [C.mid], barGapWidthPct: 55,
        showValue: true, dataLabelPosition: 'outEnd',
        dataLabelFontFace: F, dataLabelFontSize: 15, dataLabelFontBold: true, dataLabelColor: C.ink
      }
    },
    {
      type: pres.ChartType.line,
      data: [{ name: '死亡數', labels: YR, values: [19, 20, 39, 36, 36] }],
      options: {
        chartColors: [C.warn], lineSize: 3, lineSmooth: false,
        lineDataSymbol: 'circle', lineDataSymbolSize: 8, lineDataSymbolLineColor: C.warn,
        showValue: true, dataLabelPosition: 'b',
        dataLabelFontFace: F, dataLabelFontSize: 13, dataLabelFontBold: true, dataLabelColor: C.warn
      }
    }
  ], {
    x: M, y: 2.05, w: CW, h: 3.75,
    showTitle: false, showLegend: true, legendPos: 'b',
    legendFontFace: F, legendFontSize: 14, legendColor: C.ink2,
    catAxisLabelFontFace: F, catAxisLabelFontSize: 16, catAxisLabelColor: C.ink2,
    valAxisLabelFontFace: F, valAxisLabelFontSize: 13, valAxisLabelColor: C.ink2,
    valGridLine: { color: C.line, size: 1 }, catGridLine: { style: 'none' },
    valAxisMaxVal: 400
  });
  callout(s, M, 5.95, CW, 0.80,
    '五年病例數近倍增（196 → 347），同期死亡數自 2023 年起持平在 36–39 例', 'info', 18);
  foot(s, '疾管署法定傳染病統計｜查詢日 2026-08-07');
  note(s, '兩條線一起看：病例數上升是分母變大，死亡數持平代表致死率相對下降，'
    + '但這是法定傳染病通報資料，不是校正過的族群研究，不要過度解讀成「疫苗讓死亡下降」。');
}

{
  const s = slide();
  head(s, '血清型置換：疫苗涵蓋型下降', '台灣流病');
  card(s, M, 2.3, CW, 1.5, C.okBg);
  s.addText('6B　14　19F　23F　19A', {
    x: M, y: 2.3, w: CW, h: 1.5, fontFace: F, fontSize: 34, bold: true,
    color: C.ok, align: 'center', valign: 'middle', margin: 0
  });
  s.addText('幼兒常規接種產生群體免疫後，這些原本常見的血清型明顯下降', {
    x: M, y: 4.1, w: CW, h: 0.6, fontFace: F, fontSize: 21, bold: true, color: C.ink
  });
  foot(s, 'Chiang CS, et al. Vaccine 2014;32(27):3345-9｜Huang H, et al. JMII 2023;56(2):299-310');
  note(s, '先講下降的，下一頁才講上升的，對比才有力。');
}

{
  const s = slide();
  head(s, '血清型置換：非疫苗型上升', '台灣流病');
  card(s, M, 2.3, CW, 1.5, C.badBg);
  s.addText('15A　15C　23A　35B', {
    x: M, y: 2.3, w: CW, h: 1.5, fontFace: F, fontSize: 40, bold: true,
    color: C.bad, align: 'center', valign: 'middle', margin: 0
  });
  bullets(s, M, 4.1, CW, [
    '2012–2014 年 15A、15B、23A 逐年增加；2018 年起 34、35B 亦增加',
    '韓國、日本觀察到相同現象'
  ], 20);
  foot(s, 'Chiang CS, et al. Vaccine 2014;32(27):3345-9｜Huang H, et al. JMII 2023;56(2):299-310');
  note(s, '記住這四型，下一節會發現它們全都在 PCV21 裡。');
}

{
  const s = slide();
  head(s, '這是本場的核心觀察', '台灣流病');
  callout(s, M, 2.3, CW, 1.6,
    'PCV21 獨有的 8 個血清型當中，15A、15C、23A、35B 四型\n正是台灣近年上升的主流型', 'ok');
  s.addText('這是台灣本土資料，不是套用他國數字。', {
    x: M, y: 4.3, w: CW, h: 0.6, fontFace: F, fontSize: 22, bold: true, color: C.ink
  });
  callout(s, M, 5.15, CW, 0.85,
    '但現有公開資料只有趨勢方向，沒有精確百分比，因此本簡報不重算涵蓋率', 'warn');
  foot(s, '內科學誌 2025;36(6):381-386｜查詢日 2026-08-07');
  note(s, '講完這頁，聽眾就理解 PCV21 為什麼是為成人重新設計的。');
}

/* ═══════════ 三、血清型組成 ═══════════ */
{
  const s = slide({ dark: true });
  s.addText('02', { x: M, y: 2.3, w: 2.0, h: 1.0, fontFace: F, fontSize: 54, bold: true, color: '4E93A8' });
  s.addText('四支疫苗的血清型組成', {
    x: M, y: 3.3, w: CW, h: 0.9, fontFace: F, fontSize: 40, bold: true, color: 'FFFFFF'
  });
  s.addText('PCV20 與 PCV21 不是多一價的關係', {
    x: M, y: 4.2, w: CW, h: 0.6, fontFace: F, fontSize: 22, bold: true, color: 'A8CBD7'
  });
  note(s, '');
}

/* 血清型涵蓋矩陣：本節總覽，後面各頁再逐帶拆解
   32 欄 ＝ 四支疫苗血清型的聯集；分帶依 data/02「PCV20 vs PCV21 交集分析」
   欄標與格子必然小於全域 20–24pt 規範 —— 這是參照用的總覽頁，結論句仍維持大字 */
{
  const s = slide();
  head(s, '四支疫苗涵蓋的不是同一群血清型', '血清型組成');

  // 各疫苗組成（data/02，查詢日 2026-08-07）
  const SEROTYPES = {
    PPV23: ['1','2','3','4','5','6B','7F','8','9N','9V','10A','11A','12F','14','15B','17F','18C','19A','19F','20A','22F','23F','33F'],
    PCV13: ['1','3','4','5','6A','6B','7F','9V','14','18C','19A','19F','23F'],
    PCV20: ['1','3','4','5','6A','6B','7F','8','9V','10A','11A','12F','14','15B','18C','19A','19F','22F','23F','33F'],
    PCV21: ['3','6A','7F','8','9N','10A','11A','12F','15A','15C','16F','17F','19A','20A','22F','23A','23B','24F','31','33F','35B']
  };
  const VAX = ['PPV23', 'PCV13', 'PCV20', 'PCV21'].map(id => ({ id: id, color: V[id] }));
  const BANDS = [
    { cap: 'PCV20 有、PCV21 無　10 型', sub: '幼兒常規接種產生群體免疫後，\n在成人 IPD 已大幅減少',
      st: ['1','4','5','6B','9V','14','15B','18C','19F','23F'] },
    { cap: '兩者共有　10 型', sub: '選哪一支都涵蓋，\n不是選擇的考量點',
      st: ['3','6A','7F','8','10A','11A','12F','19A','22F','33F'] },
    { cap: 'PCV21 有、PCV20 無　11 型', sub: '含 15A、23A、35B 等\n成人 IPD 現行主流型',
      st: ['9N','15A','15C','16F','17F','20A','23A','23B','24F','31','35B'] },
    { cap: '', sub: '', st: ['2'] }
  ];

  // ── 建置期自我驗算：數字錯了寧可讓 build 失敗，也不要印出錯的圖 ──
  const valency = { PPV23: 23, PCV13: 13, PCV20: 20, PCV21: 21 };
  Object.keys(valency).forEach(k => {
    if (SEROTYPES[k].length !== valency[k] || new Set(SEROTYPES[k]).size !== valency[k]) {
      throw new Error('血清型矩陣：' + k + ' 價數不符或有重複');
    }
  });
  const cols = BANDS.reduce((a, b) => a.concat(b.st), []);
  const union = new Set([].concat(...Object.values(SEROTYPES)));
  if (cols.length !== union.size || new Set(cols).size !== cols.length || !cols.every(t => union.has(t))) {
    throw new Error('血清型矩陣：分帶欄位與四支疫苗聯集不一致');
  }
  const inV = (v, t) => SEROTYPES[v].indexOf(t) >= 0;
  if (!BANDS[0].st.every(t => inV('PCV20', t) && !inV('PCV21', t)) ||
      !BANDS[1].st.every(t => inV('PCV20', t) && inV('PCV21', t)) ||
      !BANDS[2].st.every(t => inV('PCV21', t) && !inV('PCV20', t)) ||
      !BANDS[3].st.every(t => inV('PPV23', t) && !inV('PCV13', t) && !inV('PCV20', t) && !inV('PCV21', t))) {
    throw new Error('血清型矩陣：分帶語意與組成表不符');
  }

  // ── 版面 ──
  const CELL = 0.30, ROWH = 0.46, BGAP = 0.18, LABW = 1.20;
  const gridW = cols.length * CELL + (BANDS.length - 1) * BGAP;
  const X0 = (W - (LABW + gridW)) / 2;
  const GX = X0 + LABW;
  const bandX = [];
  let cx = GX;
  BANDS.forEach(b => { bandX.push(cx); cx += b.st.length * CELL + BGAP; });
  const colX = t => {
    for (let b = 0; b < BANDS.length; b++) {
      const i = BANDS[b].st.indexOf(t);
      if (i >= 0) return bandX[b] + i * CELL;
    }
    return null;
  };

  // 分帶標題（第四帶僅 1 欄，說明改放註解列）
  BANDS.forEach((b, bi) => {
    if (!b.cap) return;
    const bw = b.st.length * CELL;
    s.addText(b.cap, {
      x: bandX[bi], y: 1.88, w: bw, h: 0.30, fontFace: F, fontSize: 13, bold: true,
      color: C.deep, align: 'center', valign: 'middle', margin: 0
    });
    s.addText(b.sub, {
      x: bandX[bi] - 0.1, y: 2.18, w: bw + 0.2, h: 0.46, fontFace: F, fontSize: 10.5,
      color: C.ink2, align: 'center', valign: 'top', margin: 0
    });
  });

  // 血清型欄標
  cols.forEach(t => {
    s.addText(t, {
      x: colX(t), y: 2.70, w: CELL, h: 0.26, fontFace: F, fontSize: 8.5, bold: true,
      color: C.ink2, align: 'center', valign: 'middle', margin: 0
    });
  });

  // 四列：深色＝涵蓋
  VAX.forEach((v, vi) => {
    const y = 2.98 + vi * ROWH;
    s.addText(v.id, {
      x: X0, y: y, w: LABW - 0.12, h: ROWH, fontFace: F, fontSize: 15, bold: true,
      color: v.color, align: 'right', valign: 'middle', margin: 0
    });
    cols.forEach(t => {
      s.addShape(pres.ShapeType.rect, {
        x: colX(t) + 0.015, y: y + 0.04, w: CELL - 0.03, h: ROWH - 0.08,
        fill: { color: inV(v.id, t) ? v.color : 'EDF2F5' }, line: { width: 0 }
      });
    });
  });

  s.addText('深色＝該疫苗涵蓋該血清型　│　最右欄（血清型 2）僅 PPV23 涵蓋　│　PPV23 不含 6A（仿單確認中）；20A 欄 PPV23 仿單列為 20\n'
    + 'PCV21 有、PCV20 無的 11 型中，15A、15C、16F、23A、23B、24F、31、35B 共 8 型為全新血清型，其他三支均不涵蓋', {
    x: M, y: 4.98, w: CW, h: 0.48, fontFace: F, fontSize: 10.5, color: C.ink2, valign: 'top'
  });

  callout(s, M, 5.50, CW, 1.15,
    'PCV20 比 PCV21 多 10 型、PCV21 比 PCV20 多 11 型\n兩者是不同的取捨，不是價數多寡的關係', 'warn', 20);

  foot(s, 'SAVE study. J Antimicrob Chemother 2025;80(5):1377（PMID 40131289）｜查詢日 2026-08-07');
  note(s, '這頁是本節的總覽：先讓大家看見「兩邊各多一塊、中間共有一塊」的形狀，後面三頁再把每一塊的內容念出來。'
    + '不要逐格講解，指三個分帶就好。');
}

{
  const s = slide();
  head(s, '兩者共有 10 型', 'PCV20 vs PCV21');
  card(s, M, 2.3, CW, 1.5);
  s.addText('3　6A　7F　8　10A　11A　12F　19A　22F　33F', {
    x: M, y: 2.3, w: CW, h: 1.5, fontFace: F, fontSize: 30, bold: true,
    color: V.shared, align: 'center', valign: 'middle', margin: 0
  });
  s.addText('這 10 型無論選哪一支都涵蓋，不是選擇的考量點。', {
    x: M, y: 4.2, w: CW, h: 0.6, fontFace: F, fontSize: 21, bold: true, color: C.ink
  });
  foot(s, 'SAVE study. J Antimicrob Chemother 2025;80(5):1377（PMID 40131289）｜查詢日 2026-08-07');
  note(s, '');
}

{
  const s = slide();
  head(s, 'PCV20 獨有 9 型', 'PCV20 vs PCV21');
  card(s, M, 2.3, CW, 1.4);
  s.addText('1　4　5　6B　9V　14　18C　19F　23F', {
    x: M, y: 2.3, w: CW, h: 1.4, fontFace: F, fontSize: 32, bold: true,
    color: V.PCV20, align: 'center', valign: 'middle', margin: 0
  });
  callout(s, M, 4.0, CW, 1.5,
    '這 9 型因幼兒常規接種產生群體免疫，\n在成人 IPD 已大幅減少', 'info');
  foot(s, '原文：PCV20 contains nine serotypes (excluding 15B/C) that are not contained in PCV21');
  note(s, '關鍵是「已大幅減少」——這才是 PCV21 敢捨棄它們的理由。');
}

{
  const s = slide();
  head(s, 'PCV21 獨有 11 型', 'PCV20 vs PCV21');
  card(s, M, 2.2, CW, 1.35, V.PCV21);
  s.addText('15A　15C　16F　23A　23B　24F　31　35B', {
    x: M, y: 2.2, w: CW, h: 1.35, fontFace: F, fontSize: 30, bold: true,
    color: 'FFFFFF', align: 'center', valign: 'middle', margin: 0
  });
  s.addText('以上 8 型為全新血清型，其他疫苗均不涵蓋', {
    x: M, y: 3.65, w: CW, h: 0.5, fontFace: F, fontSize: 18, bold: true, color: C.ink2, align: 'center'
  });
  card(s, M, 4.3, CW, 1.1);
  s.addText('9N　17F　20A　（PPV23 有，但 PCV20 沒有）', {
    x: M, y: 4.3, w: CW, h: 1.1, fontFace: F, fontSize: 24, bold: true,
    color: C.deep, align: 'center', valign: 'middle', margin: 0
  });
  foot(s, 'SAVE study. J Antimicrob Chemother 2025;80(5):1377（PMID 40131289）｜查詢日 2026-08-07');
  note(s, '8 ＋ 3 ＝ 11，相對 PCV20 多這 11 型。');
}

{
  const s = slide();
  head(s, '為什麼有人寫 9、有人寫 10？', '容易混淆之處');
  bullets(s, M, 2.3, CW, [
    'PCV20 含 15B，PCV21 含 15C',
    '兩者是可互相轉換的相關血清型',
    '文獻計數時通常排除此組，因此寫「9 型」'
  ], 22);
  callout(s, M, 4.7, CW, 1.1,
    '這一組差異，稍後在 STRIDE-3 試驗結果會再出現一次', 'warn');
  foot(s, 'SAVE study, J Antimicrob Chemother 2025;80(5):1377｜查詢日 2026-08-07');
  note(s, '先埋伏筆，效力那節講 15C 為什麼沒達優效時會回到這裡。');
}

{
  const s = slide();
  head(s, '兩套不同的設計哲學', '設計理念');
  card(s, M, 2.2, 5.9, 2.6);
  s.addText([
    { text: 'PCV20', options: { breakLine: true, fontSize: 26, color: V.PCV20 } },
    { text: '延續幼兒型架構\n往上加價數', options: { fontSize: 21, color: C.ink } }
  ], { x: M + 0.35, y: 2.5, w: 5.2, h: 2.0, fontFace: F, bold: true, valign: 'top' });

  card(s, 6.9, 2.2, CW - 6.2, 2.6, V.PCV21);
  s.addText([
    { text: 'PCV21', options: { breakLine: true, fontSize: 26, color: 'FFFFFF' } },
    { text: '為成人重新設計\n捨棄已消退的型、納入現行主流型', options: { fontSize: 21, color: 'E8F1F4' } }
  ], { x: 7.25, y: 2.5, w: CW - 6.9, h: 2.0, fontFace: F, bold: true, valign: 'top' });

  callout(s, M, 5.2, CW, 1.0,
    '因此四支疫苗的「涵蓋率百分比」不在同一基準上，不可直接比大小', 'bad');
  foot(s, '內科學誌 2025;36(6):381-386｜查詢日 2026-08-07');
  note(s, '這頁是整個血清型章節的結論。');
}

/* ═══════════ 四、免疫橋接教學模組（7 頁）═══════════ */
{
  const s = slide({ dark: true });
  s.addText('03', { x: M, y: 2.1, w: 2.0, h: 1.0, fontFace: F, fontSize: 54, bold: true, color: '4E93A8' });
  s.addText('免疫橋接', {
    x: M, y: 3.1, w: CW, h: 0.9, fontFace: F, fontSize: 44, bold: true, color: 'FFFFFF'
  });
  s.addText('讀懂後面所有效力數字的前提', {
    x: M, y: 4.05, w: CW, h: 0.6, fontFace: F, fontSize: 24, bold: true, color: 'A8CBD7'
  });
  s.addShape(pres.ShapeType.roundRect, {
    x: M, y: 5.0, w: 5.2, h: 0.7, rectRadius: 0.12, fill: { color: '17607A' }, line: { width: 0 }
  });
  s.addText('教學模組｜共 7 頁', {
    x: M, y: 5.0, w: 5.2, h: 0.7, fontFace: F, fontSize: 18, bold: true,
    color: 'FFFFFF', align: 'center', valign: 'middle', margin: 0
  });
  note(s, '這一節是要留給同仁帶走的觀念，講慢一點。');
}

{
  const s = slide();
  head(s, '你查不到 PCV20 與 PCV21 的「效力 XX%」', '免疫橋接 ①　破題');
  bullets(s, M, 2.4, CW, [
    'PCV13 有一個數字：45.56%',
    'PCV20、PCV21 沒有對應的數字',
    '不是資料還沒整理好——那個數字根本不存在'
  ], 23);
  callout(s, M, 5.0, CW, 1.05, '為什麼？因為它們走的是完全不同的核准路徑', 'info');
  foot(s, '');
  note(s, '用提問開場，讓聽眾自己意識到這個落差。');
}

{
  const s = slide();
  head(s, '兩種證明疫苗有效的方式', '免疫橋接 ②　對照');
  const rows = [
    [th(''), th('臨床終點試驗'), th('免疫橋接')],
    [td('看什麼', { bold: true, fill: { color: C.soft } }), td('誰生病'), td('抗體反應')],
    [td('做法', { bold: true, fill: { color: C.soft } }),
     td('隨機分兩組，追蹤數年\n數發病人數'),
     td('與已證明有效的舊疫苗\n頭對頭比 OPA 效價')],
    [td('代表', { bold: true, fill: { color: C.soft } }),
     td('PCV13', { bold: true, color: C.deep }),
     td('PCV15、PCV20、PCV21', { bold: true, color: C.bad })]
  ];
  table(s, M, 2.15, CW, rows, [2.4, (CW - 2.4) / 2, (CW - 2.4) / 2], 20, CONTENT_BOTTOM - 2.15);
  foot(s, 'OPA：調理吞噬活性，是成人試驗首選終點｜依據 V116 與 PCV20 臨床試驗計畫書（clinicaltrials.gov）｜查詢日 2026-08-07');
  note(s, 'OPA 比單純測 IgG 更接近真實殺菌功能。');
}

{
  const s = slide();
  head(s, 'CAPiTA：唯一的成人臨床終點試驗', '免疫橋接 ③　規模');
  bigStat(s, M, 2.2, 3.6, '84,496', '名 ≥65 歲成人\n荷蘭、雙盲、安慰劑對照', C.deep);
  bigStat(s, 4.6, 2.2, 3.6, '4 年', '平均追蹤時間', C.deep);
  bigStat(s, 8.5, 2.2, CW - 7.8, '45.56%', '疫苗血清型肺炎\n疫苗效力（主要終點）', C.ok);
  callout(s, M, 5.1, CW, 1.0, '要用同樣方式證明新疫苗，所需樣本數只會比 84,496 更大', 'info');
  foot(s, 'PMID 26076136；Bonten MJM, et al. N Engl J Med 2015;372:1114-25｜查詢日 2026-08-07');
  note(s, '讓聽眾感受到規模，才會理解為什麼不能一直重做。');
}

{
  const s = slide();
  head(s, '為什麼新疫苗不再做這種試驗', '免疫橋接 ③　理由');
  const items = [
    ['1', '事件太稀少', '幼兒接種後群體免疫，成人疫苗型肺炎大幅減少'],
    ['2', '倫理限制', '已有有效疫苗，不能再讓對照組打安慰劑'],
    ['3', '時間與成本', '每加一個價數就重做一次大型試驗並不可行']
  ];
  items.forEach((it, i) => {
    const y = 2.15 + i * 1.22;
    numCircle(s, M, y, 0.62, it[0]);
    s.addText(it[1], {
      x: M + 0.85, y: y - 0.04, w: 3.4, h: 0.42, fontFace: F, fontSize: 23, bold: true,
      color: C.deep, margin: 0
    });
    s.addText(it[2], {
      x: M + 0.85, y: y + 0.42, w: CW - 1.0, h: 0.48, fontFace: F, fontSize: 18,
      color: C.ink2, margin: 0
    });
  });
  callout(s, M, 5.75, CW, 0.95, '這是 FDA、EMA 都接受的法規既定路徑，不是廠商規避試驗', 'ok');
  foot(s, 'An overview of the clinical development of PCV20 in adults. Hum Vaccin Immunother 2026｜查詢日 2026-08-07');
  note(s, '強調最後一句，避免聽眾誤以為新疫苗證據薄弱。');
}

{
  const s = slide();
  head(s, 'PCV20 的橋接鏈', '免疫橋接 ④　機制');
  const rows = [
    [th('血清型'), th('橋接對象'), th('可推論的保護範圍')],
    [td('與 PCV13 共有的 13 型', { bold: true }), td('PCV13（有 CAPiTA）'), td('IPD ＋ 肺炎', { bold: true, color: C.ok })],
    [td('額外的 7 型', { bold: true }), td('PPSV23'), td('僅 IPD，不含肺炎', { bold: true, color: C.bad })]
  ];
  table(s, M, 2.15, CW, rows, [4.4, 3.8, CW - 8.2], 21, 2.2);
  callout(s, M, 4.6, CW, 1.9,
    'PPSV23 對肺炎的效果證據本身就不一致\n橋接對象證明不了的事，橋接過來也證明不了', 'bad', 23);
  foot(s, 'Hum Vaccin Immunother 2026｜查詢日 2026-08-07');
  note(s, '這是門診跟病人說「打這支預防肺炎」最容易過度延伸的地方。');
}

{
  const s = slide();
  head(s, '這套方法真正的限制', '免疫橋接 ⑤　邊界');
  callout(s, M, 2.2, CW, 1.1, '成人沒有已驗證的保護相關指標', 'bad');
  card(s, M, 3.5, CW, 1.5, C.soft);
  s.addText('"IgG antibody and OPA titer threshold values that\ncorrelate with protection in adults have not been defined."', {
    x: M + 0.4, y: 3.65, w: CW - 0.8, h: 1.2, fontFace: 'Calibri', fontSize: 19, italic: true,
    color: C.deep, valign: 'middle'
  });
  s.addText('沒有人知道 OPA 效價要多高才等於「不會發病」，\n也沒有公式能把效價差距換算成發病率差距。', {
    x: M, y: 5.2, w: CW, h: 1.0, fontFace: F, fontSize: 21, bold: true, color: C.ink
  });
  foot(s, '原文出自 V116 與 PCV20 的臨床試驗計畫書（clinicaltrials.gov）｜查詢日 2026-08-07');
  note(s, '這句話是廠商自己寫在 protocol 裡的，不是外界批評——教學時特別有說服力。');
}

{
  const s = slide();
  head(s, '換成我們熟悉的例子', '免疫橋接 ⑥　類比');
  card(s, M, 2.3, CW, 1.7, C.warnBg);
  s.addText('這就像用 LDL 降幅核准一支新的降血脂藥，\n而不是做心血管事件試驗。', {
    x: M + 0.4, y: 2.3, w: CW - 0.8, h: 1.7, fontFace: F, fontSize: 26, bold: true,
    color: C.warn, valign: 'middle'
  });
  s.addText('LDL 降得多是事實，但降多少 LDL 換到多少心肌梗塞減少，是另一個問題。', {
    x: M, y: 4.35, w: CW, h: 0.7, fontFace: F, fontSize: 21, bold: true, color: C.ink
  });
  callout(s, M, 5.35, CW, 0.9,
    '免疫橋接能回答「抗體夠不夠」，回答不了「少幾個人生病」', 'info');
  foot(s, '');
  note(s, '這是整個模組要留給聽眾的一句話。');
}

{
  const s = slide();
  head(s, '帶回門診的用語紅線', '免疫橋接 ⑦　實務');
  card(s, M, 2.2, 5.9, 3.3, C.okBg);
  s.addText('可以說', { x: M + 0.35, y: 2.4, w: 5.2, h: 0.45, fontFace: F, fontSize: 22, bold: true, color: C.ok });
  s.addText([
    { text: 'PCV21 涵蓋台灣近年上升的血清型', options: { breakLine: true } },
    { text: 'PCV21 在 11 個獨有型中的 10 個免疫原性優於 PCV20', options: { breakLine: true } },
    { text: 'PCV20／PCV21 係以免疫橋接核准', options: {} }
  ], { x: M + 0.35, y: 2.95, w: 5.2, h: 2.3, fontFace: F, fontSize: 17, bold: true, color: C.ink, paraSpaceAfter: 8, valign: 'top' });

  card(s, 6.9, 2.2, CW - 6.2, 3.3, C.badBg);
  s.addText('不可以說', { x: 7.25, y: 2.4, w: CW - 6.9, h: 0.45, fontFace: F, fontSize: 22, bold: true, color: C.bad });
  s.addText([
    { text: '「PCV21 保護力較好」', options: { breakLine: true } },
    { text: '「PCV21 較能預防肺炎」', options: { breakLine: true } },
    { text: '任何「效力 XX%」套在 PCV20／PCV21 上', options: { breakLine: true } },
    { text: '把 OPA 效價倍數講成「保護力提升 X 倍」', options: {} }
  ], { x: 7.25, y: 2.95, w: CW - 6.9, h: 2.3, fontFace: F, fontSize: 17, bold: true, color: C.ink, paraSpaceAfter: 8, valign: 'top' });

  foot(s, '免疫橋接教學模組結束');
  note(s, '這頁可以請同仁拍照帶走。');
}

/* ═══════════ 五、效力證據 ═══════════ */
{
  const s = slide({ dark: true });
  s.addText('04', { x: M, y: 2.3, w: 2.0, h: 1.0, fontFace: F, fontSize: 54, bold: true, color: '4E93A8' });
  s.addText('效力證據等級', {
    x: M, y: 3.3, w: CW, h: 0.9, fontFace: F, fontSize: 40, bold: true, color: 'FFFFFF'
  });
  s.addText('四支疫苗的「效力」不是同一種東西', {
    x: M, y: 4.2, w: CW, h: 0.6, fontFace: F, fontSize: 22, bold: true, color: 'A8CBD7'
  });
  note(s, '');
}

{
  const s = slide();
  head(s, '證據的三個層級', '效力證據');
  // 階梯高度代表「距離臨床終點的遠近」，不代表保護力高低（下方 callout 必須一起講）
  const TIER = [
    { t: '直接證據', k: '臨床終點 RCT', d: '隨機分派，以肺炎／IPD 發病為終點', v: ['PCV13'], h: 3.55 },
    { t: '間接證據', k: '觀察性研究', d: '有臨床終點，但非隨機分派', v: ['PPV23'], h: 2.95 },
    { t: '替代指標', k: '免疫橋接', d: '比 OPA 效價，不是臨床終點', v: ['PCV15', 'PCV20', 'PCV21'], h: 2.50 }
  ];
  const TW = 3.6, TGAP = 0.55, TBOT = 5.72, CHIPY = 5.00, INNER = TW - 0.44;
  TIER.forEach((t, i) => {
    const x = M + i * (TW + TGAP), y = TBOT - t.h;
    card(s, x, y, TW, t.h, C.soft);
    s.addText(t.t, {
      x: x + 0.22, y: y + 0.16, w: INNER, h: 0.34, fontFace: F, fontSize: 15, bold: true,
      color: C.ink2, margin: 0
    });
    s.addText(t.k, {
      x: x + 0.22, y: y + 0.52, w: INNER, h: 0.52, fontFace: F, fontSize: 22, bold: true,
      color: C.deep, margin: 0
    });
    s.addText(t.d, {
      x: x + 0.22, y: y + 1.06, w: INNER, h: 0.62, fontFace: F, fontSize: 14,
      color: C.ink2, valign: 'top', margin: 0
    });
    // 疫苗色票橫排並底部對齊，三欄的色票在同一水平線上
    const cw = Math.min(1.85, (INNER - (t.v.length - 1) * 0.08) / t.v.length);
    t.v.forEach((vid, j) => {
      const cx = x + 0.22 + j * (cw + 0.08);
      s.addShape(pres.ShapeType.roundRect, {
        x: cx, y: CHIPY, w: cw, h: 0.46, rectRadius: 0.10,
        fill: { color: V[vid] }, line: { width: 0 }
      });
      s.addText(vid, {
        x: cx, y: CHIPY, w: cw, h: 0.46, fontFace: F, fontSize: 15, bold: true,
        color: 'FFFFFF', align: 'center', valign: 'middle', margin: 0
      });
    });
  });
  callout(s, M, 5.90, CW, 0.85,
    '階梯講的是證據「類型」，不是保護力高低——四支之間沒有頭對頭療效試驗可以比優劣', 'bad', 18);
  foot(s, 'Hum Vaccin Immunother 2025／2026；內科學誌 2025;36(6):381-386｜查詢日 2026-08-07');
  note(s, '這頁是整節的骨架。最容易被誤讀成「PCV21 證據最弱所以比較差」——'
    + '要主動說明：新疫苗不做臨床終點試驗是因為對照組已有有效疫苗，做安慰劑對照不合倫理。');
}

{
  const s = slide();
  head(s, '證據等級一覽', '效力證據');
  const rows = [
    [th('疫苗'), th('證據類型'), th('說明')],
    [td('PPV23', { bold: true, color: V.PPV23 }), td('觀察性研究'), td('對 IPD 有保護；對肺炎證據不一致')],
    [td('PCV13', { bold: true, color: V.PCV13 }), td('臨床終點 RCT', { bold: true, color: C.ok }), td('CAPiTA，84,496 人')],
    [td('PCV15', { bold: true, color: V.PCV15 }), td('免疫橋接'), td('以安全性與免疫原性核准')],
    [td('PCV20', { bold: true, color: V.PCV20 }), td('免疫橋接'), td('13 型橋接 PCV13；7 型橋接 PPSV23')],
    [td('PCV21', { bold: true, color: V.PCV21 }), td('免疫橋接'), td('頭對頭：11 個獨有型中 10 個優效')]
  ];
  table(s, M, 2.15, CW, rows, [2.4, 3.6, CW - 6.0], 19, CONTENT_BOTTOM - 2.15);
  foot(s, '詳見各頁出處｜查詢日 2026-08-07');
  note(s, '只有 PCV13 那列是綠的，這就是重點。');
}

{
  const s = slide();
  head(s, 'CAPiTA 的三個終點', '效力證據');
  // 這是 PCV13 的資料，柱色用 PCV13 的語意色；信賴區間比照文獻圖表放在軸下方一列
  const EP = [
    { lab: 'VT-CAP（主要終點）', ve: 45.56, ci: '95.2% CI\n21.82 – 62.49' },
    { lab: '非菌血性 VT-CAP', ve: 45.00, ci: '95.2% CI\n14.21 – 65.31' },
    { lab: 'VT-IPD', ve: 75.00, ci: '95% CI\n41.43 – 90.78' }
  ];
  s.addChart(pres.ChartType.bar, [{
    name: '疫苗效力 VE', labels: EP.map(e => e.lab), values: EP.map(e => e.ve)
  }], {
    x: M, y: 1.95, w: CW, h: 3.05,
    barDir: 'col', chartColors: [V.PCV13], barGapWidthPct: 110,
    showTitle: false, showLegend: false,
    showValue: true, dataLabelPosition: 'outEnd', dataLabelFormatCode: '0.00"%"',
    dataLabelFontFace: F, dataLabelFontSize: 20, dataLabelFontBold: true, dataLabelColor: C.ink,
    catAxisLabelFontFace: F, catAxisLabelFontSize: 15, catAxisLabelColor: C.ink,
    valAxisHidden: true, valGridLine: { style: 'none' }, catGridLine: { style: 'none' },
    valAxisMaxVal: 100
  });
  // 三欄信賴區間，對齊三根柱子
  EP.forEach((e, i) => {
    s.addText(e.ci, {
      x: M + CW * (i + 0.5) / 3 - 1.7, y: 5.02, w: 3.4, h: 0.52, fontFace: F, fontSize: 13,
      color: C.ink2, align: 'center', valign: 'top', margin: 0
    });
  });
  callout(s, M, 5.68, CW, 1.05,
    '信賴區間是 95.2% 而非 95%（期中分析 alpha 消耗），引用時照抄', 'warn', 20);
  foot(s, 'Bonten MJM, et al. N Engl J Med 2015;372:1114-25（PMID 26076136）｜VT ＝ 疫苗血清型｜查詢日 2026-08-07');
  note(s, '三根柱子一起看：兩個肺炎終點都在 45% 上下，侵襲性疾病 75% 明顯更高。'
    + '95.2% 這個細節，被問到時答得出來會很加分。');
}

{
  const s = slide();
  head(s, 'CAPiTA 沒有證明的事', '效力證據');
  callout(s, M, 2.2, CW, 1.35,
    '不可以說「PCV13 可降低肺炎住院或死亡」', 'bad');
  card(s, M, 3.8, CW, 1.5, C.soft);
  s.addText('"Prevention of all-cause CAP and mortality were exploratory endpoints,\nand the study was neither designed nor powered to assess the difference."', {
    x: M + 0.4, y: 3.95, w: CW - 0.8, h: 1.2, fontFace: 'Calibri', fontSize: 17, italic: true,
    color: C.deep, valign: 'middle'
  });
  s.addText('全因肺炎與死亡率是探索性終點，試驗未設計也無檢定力去評估。', {
    x: M, y: 5.45, w: CW, h: 0.55, fontFace: F, fontSize: 21, bold: true, color: C.warn
  });
  s.addText('它證明的是「疫苗血清型」肺炎減少 45.56%，不是所有肺炎。', {
    x: M, y: 6.0, w: CW, h: 0.55, fontFace: F, fontSize: 22, bold: true, color: C.ink
  });
  foot(s, 'PMID 26076136｜查詢日 2026-08-07');
  note(s, '這是衛教時最容易誇大的一句，務必講清楚。');
}

{
  const s = slide();
  head(s, 'STRIDE-3：PCV21 對 PCV20 頭對頭', '效力證據');
  const rows = [
    [th('項目'), th('結果')],
    [td('設計'), td('隨機雙盲，主動對照組即為 PCV20', { bold: true })],
    [td('人數'), td('隨機 2,663 人，實際接種 2,656 人')],
    [td('共有 10 型'), td('第 30 天達非劣效（p<0.0001）')],
    [td('獨有 11 型'), td('其中 10 型達優效（p<0.0001）', { bold: true, color: C.ok })],
    [td('安全性'), td('6 例死亡均判定與疫苗無關；無疫苗相關嚴重不良事件')]
  ];
  table(s, M, 2.15, CW, rows, [3.4, CW - 3.4], 20, CONTENT_BOTTOM - 2.15);
  foot(s, 'Platt HL, et al. Lancet Infect Dis 2024;24(10):1141-1150（PMID 38964361）｜查詢日 2026-08-07');
  note(s, '');
}

{
  const s = slide();
  head(s, '唯一沒達到優效的那一型', '效力證據');
  card(s, M, 2.2, CW, 1.35, C.warnBg);
  s.addText('血清型 15C　（p = 0.41）', {
    x: M, y: 2.2, w: CW, h: 1.35, fontFace: F, fontSize: 36, bold: true,
    color: C.warn, align: 'center', valign: 'middle', margin: 0
  });
  bullets(s, M, 3.8, CW, [
    '新聞稿只寫「11 個中的 10 個優效」，未指名失敗者',
    '合理推測：PCV20 含 15B，與 15C 可互轉並產生交叉反應',
    '而 15C 正是台灣近年上升的血清型之一'
  ], 20);
  callout(s, M, 6.05, CW, 0.8,
    '免疫橋接的結論必須回原始論文讀，摘要與新聞稿會遺漏關鍵細節', 'bad');
  foot(s, 'PMID 38964361｜推測部分為本簡報推論，非論文明述｜查詢日 2026-08-07');
  note(s, '呼應前面血清型章節埋的 15B/15C 伏筆。');
}

/* ═══════════ 六、政策與高風險 ═══════════ */
{
  const s = slide({ dark: true });
  s.addText('05', { x: M, y: 2.3, w: 2.0, h: 1.0, fontFace: F, fontSize: 54, bold: true, color: '4E93A8' });
  s.addText('誰該打、什麼時候打', {
    x: M, y: 3.3, w: CW, h: 0.9, fontFace: F, fontSize: 40, bold: true, color: 'FFFFFF'
  });
  s.addText('一條決策鏈，三個步驟', {
    x: M, y: 4.2, w: CW, h: 0.6, fontFace: F, fontSize: 22, bold: true, color: 'A8CBD7'
  });
  note(s, '接下來十頁是同一條決策鏈，不要拆開看。');
}

{
  const s = slide();
  head(s, '病人走進診間，只問三個問題', '決策總覽');
  const steps = [
    ['1', '符不符合公費？', '看年齡、身分、適應症', C.deep],
    ['2', '過去打過什麼？', '決定要不要等、等多久', C.warn],
    ['3', '打哪一支？', 'PCV20 或 PCV21 二擇一', C.ok]
  ];
  steps.forEach((st, i) => {
    const y = 2.2 + i * 1.45;
    card(s, M, y, CW, 1.2);
    numCircle(s, M + 0.32, y + 0.28, 0.66, st[0], st[3]);
    s.addText(st[1], {
      x: M + 1.2, y: y + 0.14, w: 4.6, h: 0.5, fontFace: F, fontSize: 26, bold: true,
      color: st[3], margin: 0
    });
    s.addText(st[2], {
      x: M + 1.2, y: y + 0.64, w: CW - 1.5, h: 0.45, fontFace: F, fontSize: 18,
      color: C.ink2, margin: 0
    });
  });
  callout(s, M, 6.05, CW, 0.75, '接下來每一頁的標題都會標明它屬於哪一個步驟', 'info', 19);
  foot(s, '');
  note(s, '這頁是全場的骨架，後面每一頁都掛回這三個問題。');
}

/* 公告原表：大原則先給，後面的步驟一、步驟二再逐項拆解 */
{
  const s = slide();
  head(s, '公告原表：疫苗轉換之銜接原則', '大原則');
  const rows = [
    [th('接種對象'), th('過往接種史'), th('公費銜接疫苗')],
    [td('• 65 歲（含）以上長者\n• 55–64 歲原住民\n• 19–64 歲 IPD\n　高風險對象',
        { rowspan: 5, bold: true, color: C.deep, valign: 'middle' }),
     td('從未接種'), td('PCV20 或 PCV21', { bold: true, color: C.ok })],
    [td('僅接種 PPV23'), td('≥ 1 年　→　PCV20 或 PCV21', { bold: true })],
    [td('僅接種 PCV13/15'), td('≥ 1 年 *　→　PCV20 或 PCV21', { bold: true })],
    [td('PCV13/15 ＋ PPV23'), td('已完整接種，無需再接種', { bold: true, color: C.ink2 })],
    [td('PCV20 或 PCV21'), td('已完整接種，無需再接種', { bold: true, color: C.ink2 })]
  ];
  table(s, M, 2.1, CW, rows, [3.1, 3.5, CW - 6.6], 17, 3.35);
  callout(s, M, 5.6, CW, 1.3,
    '* IPD 高風險對象、65 歲以上機構住民及洗腎患者為 ≥ 8 週\n' +
    '註：19–64 歲 IPD 高風險對象，如於 65 歲前完整接種 PCV13/15＋PPV23，' +
    '可於滿 65 歲（含）且與前劑間隔滿 5 年後，再追加接種 1 劑', 'warn', 17);
  foot(s, '成人肺炎鏈球菌疫苗接種須知（自 115 年 8 月 10 日起適用）伍、疫苗轉換之銜接原則｜查詢日 2026-08-08');
  note(s, '這是公告原表，是整節的權威依據。先讓大家看過原表掌握大原則，後面的步驟一、步驟二再逐項拆解。');
}

/* 決策流程圖：把上一頁的公告原表畫成一條流程
   內容全部取自 rules.json（fundedCategories / officialFlow / shortIntervalEligibility /
   vaccines），與 HTML 報告的決策器共用同一份規則表 —— 不得在此另寫一套判斷 */
{
  const s = slide();
  head(s, '同一份規則，畫成一條流程', '決策總覽');

  const G = RULES.fundedCategories;
  const FLOW = RULES.officialFlow.groups;
  const two = FLOW.find(g => g.id === 'two').subs;
  const SHORT = RULES.shortIntervalEligibility;

  const AX = M, AW = 2.45;                    // 步驟 1
  const HX = 3.45, HW = 3.05;                 // 步驟 2　接種史
  const OX = 6.95, OW = 3.05;                 // 步驟 2　銜接規則
  const CX = 10.5, CW3 = 2.10;                // 步驟 3
  const Y0 = 2.28, ROWH = 0.60, GAP = 0.06;
  const rowY = i => Y0 + i * (ROWH + GAP);

  const colHead = (x, w, t) => s.addText(t, {
    x: x, y: 1.92, w: w, h: 0.30, fontFace: F, fontSize: 13, bold: true,
    color: C.deep, align: 'center', valign: 'middle', margin: 0
  });
  colHead(AX, AW, '步驟 1　公費資格');
  colHead(HX, OX + OW - HX, '步驟 2　過去打過什麼');
  colHead(CX, CW3, '步驟 3');

  // ── 步驟 1：三類公費對象（rules.fundedCategories）──
  const gLabel = { elder65: '65 歲（含）以上', indigenous55: '55–64 歲原住民',
                   highrisk19: '19–64 歲\nIPD 高風險（5 項）' };
  G.forEach((g, i) => {
    const y = 2.67 + i * 1.16;
    card(s, AX, y, AW, 0.92);
    s.addText(gLabel[g.id], {
      x: AX + 0.12, y: y, w: AW - 0.24, h: 0.92, fontFace: F, fontSize: 15, bold: true,
      color: C.deep, align: 'center', valign: 'middle', margin: 0
    });
  });
  s.addText('符合任一項', {
    x: AX, y: 5.95, w: AW, h: 0.28, fontFace: F, fontSize: 12, bold: true,
    color: C.ink2, align: 'center', valign: 'middle', margin: 0
  });
  s.addShape(pres.ShapeType.rightArrow, {
    x: HX - 0.30, y: 4.18, w: 0.26, h: 0.22, fill: { color: C.mid }, line: { width: 0 }
  });

  // ── 步驟 2：接種史 → 銜接規則（rules.officialFlow）──
  const short = SHORT.conditions.join('、');
  const rows = [
    { h: FLOW.find(g => g.id === 'one').detail.replace('（包含 13、15、20、21 或 23 價疫苗）', ''),
      o: FLOW.find(g => g.id === 'one').action, tone: 'ok', go: true },
    { h: two[0].label.replace('（一）', ''), o: two[0].action, tone: 'info', go: true },
    { h: two[1].label.replace('（二）', ''), o: two[1].action, tone: 'info', go: true },
    { h: '└ 其中 IPD 高風險、機構住民或洗腎患者', o: '間隔至少 ' + SHORT.label + '\n（登錄加註 ' + SHORT.registrationCode.code + '）',
      tone: 'warn', go: true, sub: true },
    { h: 'PCV13/15 ＋ PPV23 完整兩劑',
      o: '三條件全符合 → 追加 1 劑\n否則視為已完成',
      tone: 'warn', go: true },
    { h: '已打過 PCV20 或 PCV21', o: '視為已完成，不再追加', tone: 'stop', go: false }
  ];
  const TONE = { ok: [C.okBg, C.ok], info: [C.soft, C.deep], warn: [C.warnBg, C.warn], stop: ['F0F3F5', C.ink2] };

  rows.forEach((r, i) => {
    const y = rowY(i);
    const pair = TONE[r.tone];
    card(s, HX, y, HW, ROWH, r.sub ? 'F7FAFB' : C.soft);
    s.addText(r.h, {
      x: HX + (r.sub ? 0.20 : 0.14), y: y, w: HW - 0.34, h: ROWH, fontFace: F,
      fontSize: r.sub ? 11.5 : 13, bold: !r.sub, color: r.sub ? C.ink2 : C.ink,
      valign: 'middle', margin: 0
    });
    card(s, OX, y, OW, ROWH, pair[0]);
    s.addText(r.o, {
      x: OX + 0.14, y: y, w: OW - 0.28, h: ROWH, fontFace: F, fontSize: 12.5, bold: true,
      color: pair[1], valign: 'middle', margin: 0
    });
    if (r.go) {
      s.addShape(pres.ShapeType.rightArrow, {
        x: OX + OW + 0.10, y: y + ROWH / 2 - 0.10, w: 0.26, h: 0.20,
        fill: { color: C.ok }, line: { width: 0 }
      });
    }
  });
  s.addText('8 週適用對象：' + short + '　│　'
    + '「三條件」＝先前為 19–64 歲 IPD 高風險 ＋ 滿 65 歲（含）＋ 與前劑間隔滿 5 年，缺一不可', {
    x: M, y: rowY(rows.length - 1) + ROWH + 0.10, w: CW, h: 0.44, fontFace: F,
    fontSize: 10, color: C.ink2, valign: 'top'
  });

  // ── 步驟 3：兩支擇一（rules.vaccines）──
  const goRows = rows.filter(r => r.go).length;
  const spanTop = rowY(0), spanBot = rowY(goRows - 1) + ROWH;
  card(s, CX, (spanTop + spanBot) / 2 - 0.95, CW3, 1.60, C.okBg);
  s.addText(RULES.vaccines.offered.join('\n或\n'), {
    x: CX, y: (spanTop + spanBot) / 2 - 0.95, w: CW3, h: 1.10, fontFace: F, fontSize: 17,
    bold: true, color: C.ok, align: 'center', valign: 'middle', margin: 0
  });
  s.addText('擇一接種 1 劑', {
    x: CX, y: (spanTop + spanBot) / 2 + 0.05, w: CW3, h: 0.55, fontFace: F, fontSize: 13,
    bold: true, color: C.ink, align: 'center', valign: 'middle', margin: 0
  });
  s.addText(RULES.vaccines.choiceNote, {
    x: CX - 0.05, y: (spanTop + spanBot) / 2 + 0.78, w: CW3 + 0.1, h: 0.80, fontFace: F,
    fontSize: 10.5, color: C.ink2, align: 'center', valign: 'top', margin: 0
  });

  foot(s, '公告「伍、疫苗轉換之銜接原則」｜本圖由 rules.json 產生，與 HTML 決策器同源｜查詢日 2026-08-08');
  note(s, '上一頁是公告原表，這一頁是同一份規則的流程版。講法：左邊三類決定「有沒有公費」，'
    + '中間五種接種史決定「要等多久」，右邊永遠是同一個結論。最後一列（已打過 PCV20/PCV21）沒有箭頭，代表流程到此結束。'
    + '第四列的 8 週是第三列的但書，不是獨立的一類。');
}

{
  const s = slide();
  head(s, '19–64 歲公費高風險：僅這 5 項', '步驟 1　公費資格');
  const items = [
    '脾臟功能缺損',
    '先天或後天免疫功能不全',
    '人工耳植入',
    '腦脊髓液滲漏',
    '「一年內」接受免疫抑制劑或放射治療的惡性腫瘤者及器官移植者'
  ];
  items.forEach((t, i) => {
    const y = 2.1 + i * 0.66;
    card(s, M, y, CW, 0.58, C.okBg);
    numCircle(s, M + 0.18, y + 0.06, 0.46, String(i + 1), C.ok);
    s.addText(t, {
      x: M + 0.8, y: y, w: CW - 1.0, h: 0.58, fontFace: F, fontSize: 20, bold: true,
      color: C.ink, valign: 'middle', margin: 0
    });
  });
  callout(s, M, 5.5, CW, 1.2,
    '這 5 項決定的是「公費對象資格」。\n' +
    '三類公費對象之後適用完全相同的銜接原則，不因身分別而異。', 'info', 19);
  foot(s, '成人肺炎鏈球菌疫苗接種須知（自 115 年 8 月 10 日起適用）參、實施對象｜查詢日 2026-08-08');
  note(s, '重點只有一個：清單就是這 5 項，糖尿病與 COPD 不在其中。公告原表的「接種對象」欄是跨列合併的，三類對象套同一套銜接原則，不要讓聽眾誤以為高風險族群另有一套較嚴的規則。');
}

{
  const s = slide();
  head(s, '高風險必須分兩層看', '步驟 1　公費資格');
  const rows = [
    [th('層'), th('適應症'), th('19–64 歲'), th('65 歲以上')],
    [td('第一層\n公費高風險', { bold: true }),
     td('脾臟功能缺損、免疫功能不全、人工耳植入、腦脊髓液滲漏、惡性腫瘤、器官移植'),
     td('公費', { bold: true, color: C.ok }), td('公費', { bold: true, color: C.ok })],
    [td('第二層\n臨床高風險\n但非公費', { bold: true }),
     td('糖尿病、慢性心臟病、COPD、氣喘、慢性肝病、慢性腎病、酒精使用障礙、吸菸'),
     td('自費', { bold: true, color: C.bad }), td('公費', { bold: true, color: C.ok })]
  ];
  table(s, M, 2.15, CW, rows, [2.7, 5.9, 1.7, CW - 10.3], 19, CONTENT_BOTTOM - 2.15);
  foot(s, '疾管署致醫界通函第 612 號｜查詢日 2026-08-07');
  note(s, '不能籠統說「高風險就公費」或「高風險只能自費」，兩種說法都會給錯答案。');
}

{
  const s = slide();
  head(s, '分界是年齡，不是疾病', '步驟 1　公費資格');
  card(s, M, 2.3, 5.9, 2.2, C.badBg);
  s.addText([
    { text: '62 歲　糖尿病', options: { breakLine: true, fontSize: 28, color: C.ink } },
    { text: '自　費', options: { fontSize: 34, color: C.bad } }
  ], { x: M, y: 2.3, w: 5.9, h: 2.2, fontFace: F, bold: true, align: 'center', valign: 'middle', margin: 0 });

  card(s, 6.9, 2.3, CW - 6.2, 2.2, C.okBg);
  s.addText([
    { text: '65 歲　同一個人', options: { breakLine: true, fontSize: 28, color: C.ink } },
    { text: '公　費', options: { fontSize: 34, color: C.ok } }
  ], { x: 6.9, y: 2.3, w: CW - 6.2, h: 2.2, fontFace: F, bold: true, align: 'center', valign: 'middle', margin: 0 });

  callout(s, M, 4.95, CW, 1.35,
    '文獻：具 2 個以上共病症者，肺炎風險與高風險患者相當或甚至更高\n對這群人應主動評估是否自費接種', 'warn');
  foot(s, 'Pelton SI, et al. BMC Infect Dis 2015;15:470｜查詢日 2026-08-07');
  note(s, '這一頁是整場最實用的一頁。');
}

/* 步驟 2　銜接規則：逐項拆解（公告原表已移到決策總覽之後先給） */
{
  const s = slide();
  head(s, '公告把接種資格分成兩大類', '步驟 2　過去打過什麼');
  const cw2 = 5.75;

  card(s, M, 2.15, cw2, 4.4, C.okBg);
  s.addText('對象一', { x: M + 0.35, y: 2.4, w: cw2 - 0.7, h: 0.5, fontFace: F, fontSize: 26, bold: true, color: C.ok });
  s.addText('從未接種', { x: M + 0.35, y: 2.95, w: cw2 - 0.7, h: 0.55, fontFace: F, fontSize: 30, bold: true, color: C.ink });
  s.addText('包含 13、15、20、21 或 23 價疫苗', {
    x: M + 0.35, y: 3.6, w: cw2 - 0.7, h: 0.5, fontFace: F, fontSize: 17, color: C.ink2 });
  s.addText('無間隔需求\n即可接種 1 劑', {
    x: M + 0.35, y: 4.4, w: cw2 - 0.7, h: 1.7, fontFace: F, fontSize: 28, bold: true,
    color: C.ok, valign: 'middle' });

  card(s, M + cw2 + 0.4, 2.15, cw2, 4.4);
  const x2 = M + cw2 + 0.75;
  s.addText('對象二', { x: x2, y: 2.4, w: cw2 - 0.7, h: 0.5, fontFace: F, fontSize: 26, bold: true, color: C.deep });
  s.addText('曾經接種', { x: x2, y: 2.95, w: cw2 - 0.7, h: 0.55, fontFace: F, fontSize: 30, bold: true, color: C.ink });
  s.addText('依下列三種情形銜接接種', {
    x: x2, y: 3.6, w: cw2 - 0.7, h: 0.5, fontFace: F, fontSize: 17, color: C.ink2 });
  s.addText([
    { text: '（一）僅接種過 23 價', options: { breakLine: true } },
    { text: '（二）僅接種過 13 或 15 價', options: { breakLine: true } },
    { text: '（三）曾完整接種兩劑', options: {} }
  ], { x: x2, y: 4.35, w: cw2 - 0.7, h: 1.8, fontFace: F, fontSize: 21, bold: true,
       color: C.deep, paraSpaceAfter: 8, valign: 'top' });

  foot(s, '疾管署 2026-08-10 起實施之公費接種公告｜合約醫療院所約 2,800 家｜查詢日 2026-08-08');
  note(s, '先把公告的兩層結構講清楚，後面三頁才逐項展開對象二。');
}

{
  const s = slide();
  head(s, '對象二（一）（二）：只打過一種疫苗', '步驟 2　過去打過什麼');
  const rows = [
    [th('情形'), th('過去接種史'), th('間隔要求')],
    [td('（一）', { bold: true, color: C.deep }), td('僅接種過 23 價疫苗'), td('≥ 1 年', { bold: true })],
    [td('（二）', { bold: true, color: C.deep }), td('僅接種過 13 或 15 價疫苗'), td('≥ 1 年', { bold: true })],
    [td('（二）\n但書', { bold: true, color: C.bad }),
     td('上列且為 IPD 高風險對象，或 65 歲以上機構住民、洗腎患者'),
     td('≥ 8 週', { bold: true, color: C.bad })]
  ];
  table(s, M, 2.15, CW, rows, [1.9, 7.0, CW - 8.9], 20, CONTENT_BOTTOM - 2.15);
  foot(s, '疾管署 2026-08-10 起實施之公費接種公告｜查詢日 2026-08-08');
  note(s, '最後一列的但書最常被漏掉——高風險對象本身就適用 8 週，不必等一年。');
}

{
  const s = slide();
  head(s, '（二）的但書：誰適用 8 週', '步驟 2　過去打過什麼');
  const items = RULES.shortIntervalEligibility.conditions;
  items.forEach((t, i) => {
    const y = 2.3 + i * 1.25;
    card(s, M, y, CW, 1.05, C.badBg);
    numCircle(s, M + 0.3, y + 0.2, 0.65, String(i + 1), C.bad);
    s.addText(t, {
      x: M + 1.15, y: y, w: CW - 1.5, h: 1.05, fontFace: F, fontSize: 26, bold: true,
      color: C.ink, valign: 'middle', margin: 0
    });
  });
  callout(s, M, 6.1, CW, 0.75,
    '限曾接種 13 或 15 價者；其餘情形仍為 1 年　｜　登錄時加註身分別代碼「'
    + RULES.shortIntervalEligibility.registrationCode.code + '」', 'info', 19);
  foot(s, '疾管署 2026-08-10 起實施之公費接種公告｜查詢日 2026-08-08');
  note(s, '這三種身分只要打過 13 或 15 價，八週後就可以接種新疫苗。注意第一項「IPD 高風險對象」公告未限年齡，65 歲以上的高風險者同樣適用 8 週。登錄時記得加註 R02A，統計才分得出來。');
}

{
  const s = slide();
  head(s, '對象二（三）：三個條件缺一不可', '步驟 2　過去打過什麼');
  s.addText('已完整接種 13 價（或 15 價）＋ 23 價各 1 劑者，要再接種 1 劑必須同時符合：', {
    x: M, y: 1.95, w: CW, h: 0.5, fontFace: F, fontSize: 20, bold: true, color: C.ink2 });

  const cw3 = 3.6, gap = 0.55;
  const conds = [
    ['1', '先前為\n19–64 歲\nIPD 高風險對象', C.deep],
    ['2', '已滿\n65 歲（含）', C.bad],
    ['3', '與前劑\n間隔滿 5 年', C.deep]
  ];
  conds.forEach((cd, i) => {
    const x = M + i * (cw3 + gap);
    card(s, x, 2.55, cw3, 2.85, i === 1 ? C.badBg : C.soft);
    numCircle(s, x + cw3 / 2 - 0.33, 2.78, 0.66, cd[0], cd[2]);
    s.addText(cd[1], {
      x: x + 0.2, y: 3.55, w: cw3 - 0.4, h: 1.65, fontFace: F, fontSize: 23, bold: true,
      color: cd[2], align: 'center', valign: 'top', margin: 0
    });
    if (i < 2) {
      s.addText('＋', {
        x: x + cw3 - 0.2, y: 3.5, w: gap + 0.4, h: 0.6, fontFace: F, fontSize: 26, bold: true,
        color: C.ink2, align: 'center', valign: 'middle', margin: 0
      });
    }
  });

  callout(s, M, 5.65, CW, 1.25,
    '三項是「同時成立」，缺一即視為已完成接種、不再追加。\n最容易漏掉的是第 2 項：即使先前是 IPD 高風險對象，未滿 65 歲仍不可追加。', 'bad', 20);
  foot(s, '疾管署 2026-08-10 起實施之公費接種公告｜查詢日 2026-08-08');
  note(s, '這一條最容易讀成「任何 65 歲以上打過兩劑都可以再打」，或「高風險就可以再打」。三個條件是 AND，不是 OR，而且 65 歲那條沒有例外。');
}

/* 步驟 3 */
{
  const s = slide();
  head(s, '打哪一支？', '步驟 3　選擇疫苗');
  card(s, M, 2.3, CW, 1.9, C.okBg);
  s.addText('PCV20　或　PCV21', {
    x: M, y: 2.3, w: CW, h: 1.9, fontFace: F, fontSize: 46, bold: true,
    color: C.ok, align: 'center', valign: 'middle', margin: 0 });
  callout(s, M, 4.5, CW, 1.0,
    '疾管署表示兩者無優劣差異，由民眾自由選擇其一接種', 'info', 22);
  callout(s, M, 5.7, CW, 1.15,
    '若病人問「差在哪」：PCV21 涵蓋 15A、15C、23A、35B 等台灣近年上升的血清型；\nPCV20 涵蓋 1、4、5、6B 等已因幼兒接種而大幅減少的血清型', 'warn', 18);
  foot(s, '疾管署 2026-08-10 起實施之公費接種公告｜查詢日 2026-08-08');
  note(s, '官方立場是無優劣、自由選擇。要補充差異時只講血清型組成，不要講保護力高低。');
}

/* 補充：台美差異（不屬於決策鏈，放在三步驟之後） */
{
  const s = slide();
  head(s, '台美差異速查', '補充｜非決策步驟');
  const rows = [
    [th('項目'), th('台灣公費'), th('美國 ACIP')],
    [td('起始年齡'), td('65 歲', { bold: true, color: C.bad }), td('50 歲', { bold: true, color: C.ok })],
    [td('高風險定義'), td('窄：僅 5 項'), td('寬：另含糖尿病、慢性心肺肝腎病、吸菸')],
    [td('PPV23'), td('不再提供公費', { bold: true }), td('仍為方案二的一部分')],
    [td('50–64 歲一般成人'), td('不符公費', { bold: true, color: C.bad }), td('建議接種', { bold: true, color: C.ok })]
  ];
  table(s, M, 2.15, CW, rows, [2.9, 3.6, CW - 6.5], 19, [0.62, 0.62, 1.05, 0.62, 0.62]);
  callout(s, M, 5.85, CW, 1.05, '門診最常見的落差：50–64 歲有糖尿病或 COPD 的病人', 'warn', 22);
  foot(s, 'ACIP 為國際指引，非疾管署公告內容｜MMWR Recomm Rep 2023;72(3):1-39；PMID 39773952');
  note(s, '這頁是比較，不是台灣的規定。決定公費與否一律看前面三個步驟。');
}

/* ═══════════ 六、合併施打 ═══════════ */
{
  const s = slide({ dark: true });
  s.addText('06', { x: M, y: 2.3, w: 2.0, h: 1.0, fontFace: F, fontSize: 54, bold: true, color: '4E93A8' });
  s.addText('與其他疫苗的合併施打', {
    x: M, y: 3.3, w: CW, h: 0.9, fontFace: F, fontSize: 40, bold: true, color: 'FFFFFF'
  });
  s.addText('九種疫苗、三個例外', {
    x: M, y: 4.2, w: CW, h: 0.6, fontFace: F, fontSize: 22, bold: true, color: 'A8CBD7'
  });
  note(s, '');
}

{
  const s = slide();
  head(s, '合併施打：一句話原則', '門診實務');
  callout(s, M, 2.3, CW, 1.7,
    '肺炎鏈球菌疫苗全部都是不活化疫苗\n與任何疫苗都可同日於不同部位接種，不需要任何間隔', 'ok');
  callout(s, M, 4.35, CW, 1.5,
    '本節非疾管署公告內容\n通函第 612 號未規範合併施打，以下依國際指引與各疫苗研究整理', 'warn');
  callout(s, M, 6.0, CW, 0.72,
    '如疾管署日後發布相關規定，本頁內容依公告修訂', 'info', 19);
  foot(s, 'CDC Yellow Book 2026（出版日 2025-04-23）｜查詢日 2026-08-07');
  note(s, '一定要講「非官方公告」這句，避免聽眾以為有疾管署背書。');
}

{
  const s = slide();
  head(s, '九種疫苗，同一個答案', '門診實務');
  s.addText('整欄都是「可同日、不需間隔」——因為肺炎鏈球菌疫苗全部都是不活化疫苗', {
    x: M, y: 1.82, w: CW, h: 0.34, fontFace: F, fontSize: 16, bold: true, color: C.ink2
  });
  const yes = t => td(t, { bold: true, color: C.ok, align: 'center' });
  // 規格卡要求「不活化與活性減毒分兩區呈現」，用 rowspan 做出分區（與 HTML 報告同一種呈現）
  const GRP = [
    { g: '不活化\n／重組', rows: [
      ['流感', '長年併用，最常見組合'],
      ['COVID-19', '可與幾乎所有疫苗同時接種'],
      ['RSV', '⚠ 併打時 RSV 與流感效價略低，不構成禁忌'],
      ['Tdap', '與 RZV 併用研究無免疫干擾'],
      ['A 型肝炎', '不同部位、不同針筒'],
      ['B 型肝炎', '不同部位、不同針筒'],
      ['帶狀疱疹 Shingrix', '重組佐劑型，不是活性疫苗，常被搞錯']
    ] },
    { g: '活性\n減毒', rows: [
      ['MMR', '⚠ 免疫功能不全者禁用（與肺鏈同日無關）'],
      ['水痘', '⚠ 免疫功能不全者禁用（與肺鏈同日無關）']
    ] }
  ];
  const rows = [[th('分類'), th('疫苗'), th('與肺鏈同日'), th('間隔'), th('備註')]];
  GRP.forEach(grp => {
    grp.rows.forEach((r, i) => {
      const line = [];
      if (i === 0) line.push(td(grp.g, { rowspan: grp.rows.length, bold: true, color: C.deep, valign: 'middle' }));
      line.push(td(r[0], { bold: true }), yes('✔ 可'),
                td('不需', { align: 'center', color: C.ink2 }), td(r[1]));
      rows.push(line);
    });
  });
  table(s, M, 2.22, CW, rows, [1.35, 2.45, 1.5, 1.2, CW - 6.5], 14, CONTENT_BOTTOM - 2.22);
  foot(s, 'CDC Yellow Book 2026（出版日 2025-04-23）｜SHINGRIX coadministration（GSK）｜非疾管署公告內容｜查詢日 2026-08-07');
  note(s, '重點是「整欄一樣」——不要一列一列念。只要指出兩件事：Shingrix 不是活性疫苗；'
    + 'MMR／水痘的禁忌是針對免疫功能不全者本身，跟肺鏈同日與否無關。三個真正的例外在下一頁。'
    + '被問到帶狀疱疹時可補充：台灣已無活性減毒型（Zostavax 於 2025-10-07 自請註銷），'
    + '現行只有 Shingrix，所以帶疹疫苗不會落到活性減毒那一區。');
}

{
  const s = slide();
  head(s, '最容易寫錯的一條', '門診實務');
  callout(s, M, 2.2, CW, 1.5,
    '「間隔 ≥28 天」規範的是兩種活性減毒疫苗彼此之間\n不是肺鏈與活性疫苗之間', 'bad');
  bullets(s, M, 4.0, CW, [
    '肺鏈是不活化疫苗，與 MMR／水痘同日或任何間隔皆可',
    '只有 MMR 與水痘互相之間未同日接種，才須間隔 ≥28 天',
    '間隔不足 28 天者，第二劑視為無效，須再隔 ≥28 天重打'
  ], 20);
  foot(s, 'CDC Yellow Book 2026 逐字：injectable or nasally administered live vaccines…at intervals of ≥28 days');
  note(s, '我自己一開始也記錯，這條值得特別提醒同仁。');
}

{
  const s = slide();
  head(s, '三個例外', '門診實務');
  const items = [
    ['1', '無脾症或 HIV 感染者', 'PCV13 與 MenACWY-D（Menactra）不可同時接種；先打 PCV13，至少 4 週後再打。台灣無 Menactra 藥證，現行的 Menveo（MenACWY-CRM）無此限制，本條只適用有國外接種史者'],
    ['2', 'PCV15 與 PPV23', '不可於同一次就診接種'],
    ['3', 'RSV 疫苗', '併打時效價略低，可考慮分開，但不構成禁忌']
  ];
  items.forEach((it, i) => {
    const y = 2.25 + i * 1.42;
    numCircle(s, M, y, 0.6, it[0], C.warn);
    s.addText(it[1], {
      x: M + 0.82, y: y - 0.05, w: CW - 1.0, h: 0.42, fontFace: F, fontSize: 22, bold: true,
      color: C.deep, margin: 0
    });
    s.addText(it[2], {
      x: M + 0.82, y: y + 0.38, w: CW - 1.0, h: 0.85, fontFace: F, fontSize: 16,
      color: C.ink2, margin: 0
    });
  });
  foot(s, 'CDC Pink Book Ch.2、Ch.14｜非疾管署公告內容｜第 1 條僅針對 PCV13，未延伸至 PCV20／21｜藥證：食藥署開放資料｜查詢日 2026-08-09');
  note(s, '第一條原本直接命中公費高風險族群，但台灣查無 Menactra 藥證，'
       + '現行四價流腦疫苗是 Menveo（腦寧安，MenACWY-CRM），不受此限。'
       + '真正會遇到的情境是病人在國外打過 Menactra。');
}

/* ═══════════ 八、結語 ═══════════ */
{
  const s = slide();
  head(s, '同時提供的工具', '收尾');
  card(s, M, 2.3, CW, 2.5);
  s.addText([
    { text: '互動式 HTML 報告', options: { breakLine: true, fontSize: 26, color: C.deep } },
    { text: '內建接種建議決策器：輸入西元出生年、接種史、間隔與適應症，', options: { breakLine: true, fontSize: 19, color: C.ink } },
    { text: '直接輸出「該不該打 → 打哪支 → 公費或自費」，並附台美建議對照。', options: { fontSize: 19, color: C.ink } }
  ], { x: M + 0.4, y: 2.55, w: CW - 0.8, h: 2.0, fontFace: F, bold: true, valign: 'top' });
  callout(s, M, 5.05, CW, 1.0, '單一檔案、可離線開啟，每個數字都可回溯到原始出處', 'info');
  foot(s, '');
  note(s, '報告與本簡報共用同一份規則表，數字保證一致。');
}

{
  const s = slide({ dark: true });
  s.addText('三個帶走的重點', {
    x: M, y: 1.3, w: CW, h: 0.8, fontFace: F, fontSize: 36, bold: true, color: 'FFFFFF'
  });
  const pts = [
    ['1', '新制 8/10 上路', 'PCV20 或 PCV21 一劑搞定，PPV23 不再公費'],
    ['2', '高風險分兩層', '19–64 歲僅 5 項公費；糖尿病、COPD 要自費'],
    ['3', '效力數字要看證據等級', '新疫苗是免疫橋接核准，不可講成保護力較好']
  ];
  pts.forEach((p, i) => {
    const y = 2.5 + i * 1.35;
    numCircle(s, M, y, 0.66, p[0], '4E93A8');
    s.addText(p[1], {
      x: M + 0.9, y: y - 0.05, w: CW - 1.0, h: 0.45, fontFace: F, fontSize: 24, bold: true,
      color: 'FFFFFF', margin: 0
    });
    s.addText(p[2], {
      x: M + 0.9, y: y + 0.42, w: CW - 1.0, h: 0.45, fontFace: F, fontSize: 18,
      color: 'A8CBD7', margin: 0
    });
  });
  note(s, '結尾只留三句，不要再加新資訊。');
}

{
  const s = slide();
  head(s, '本簡報的來源位階原則', '附錄');
  callout(s, M, 2.05, CW, 1.0,
    '政策陳述一律以台灣官方公告為準：疾管署、衛福部、健保署', 'bad', 23);
  const rows = [
    [th('位階'), th('來源'), th('用途')],
    [td('1', { bold: true, color: C.deep }), td('台灣官方公告\n致醫界通函、衛福部新聞稿、接種須知', { bold: true }),
     td('公費資格、間隔規定、適應症認定——只引這一層', { bold: true, color: C.ok })],
    [td('2', { bold: true, color: C.deep }), td('台灣學會共識與本土期刊'), td('臨床判斷與本土流病，非政策依據')],
    [td('3', { bold: true, color: C.deep }), td('國際指引與期刊\nACIP、CDC Yellow Book、Lancet'), td('臨床證據；須標明「非疾管署公告內容」')],
    [td('4', { bold: true, color: C.bad }), td('藥廠新聞稿、媒體、衛教網站'), td('不得作為政策或效力陳述的依據', { bold: true, color: C.bad })]
  ];
  table(s, M, 3.25, CW, rows, [1.1, 4.6, CW - 5.7], 16, 3.3);
  foot(s, '官方文件互有出入時，以給醫事人員的正式函文為準並標示差異｜本主題主管機關為疾管署');
  note(s, '這一頁是方法論宣告：政策的事只信官方公告，國際指引一律標示非官方。');
}

{
  const s = slide();
  head(s, '主要出處', '附錄');
  const rows = [
    [th('內容'), th('出處')],
    [td('公費政策與銜接規則'), td('疾管署致醫界通函第 612 號（2026-07-28）')],
    [td('台灣流病與血清型置換'), td('內科學誌 2025;36(6):381-386')],
    [td('血清型組成'), td('J Antimicrob Chemother 2025;80(5):1377（PMID 40131289）')],
    [td('PCV13 效力'), td('CAPiTA，PMID 26076136；NEJM 2015;372:1114-25')],
    [td('PCV21 頭對頭'), td('Lancet Infect Dis 2024;24(10):1141-1150（PMID 38964361）')],
    [td('ACIP 建議'), td('MMWR Recomm Rep 2023;72(3):1-39；PMID 39773952')],
    [td('合併施打'), td('CDC Yellow Book 2026（2025-04-23）')]
  ];
  table(s, M, 2.1, CW, rows, [3.8, CW - 3.8], 17, CONTENT_BOTTOM - 2.1);
  foot(s, '全部查證日 2026-08-07｜本簡報未使用任何病人資料');
  note(s, '');
}

/* ─────────── 輸出 ─────────── */
const OUT_DIR = path.join(__dirname, '..', 'output');
const OUT = process.env.DECK_OUT || path.join(OUT_DIR, '台灣成人肺炎鏈球菌疫苗簡報.pptx');
/* 備援檔名帶時間戳。
   （曾叫「_新版」，結果那個檔停在舊版本卻頂著「新版」的名字，反過來誤導。） */
function stamp() {
  const d = new Date(), p = x => String(x).padStart(2, '0');
  return d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate()) + '-' + p(d.getHours()) + p(d.getMinutes());
}
const FALLBACK = path.join(OUT_DIR, '台灣成人肺炎鏈球菌疫苗簡報_備援_' + stamp() + '.pptx');
fs.mkdirSync(OUT_DIR, { recursive: true });

pres.writeFile({ fileName: OUT })
  .then(() => {
    console.log('✅ 產生完成：' + OUT);
    console.log('   投影片張數：' + n);
  })
  .catch((err) => {
    if (err.code !== 'EBUSY') throw err;
    // 正式檔被 PowerPoint 開著時不硬蓋，改寫備用檔，避免整個建置中斷
    return pres.writeFile({ fileName: FALLBACK }).then(() => {
      console.log('⚠️  正式檔被開啟中（EBUSY），未覆蓋。');
      console.log('   已改寫備用檔：' + FALLBACK);
      console.log('   投影片張數：' + n);
      console.log('   關閉 PowerPoint 後重跑 node src/deck.js 即可覆蓋正式檔。');
    });
  });
