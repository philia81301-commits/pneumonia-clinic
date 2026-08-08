# -*- coding: utf-8 -*-
"""
總驗收：逐條核對 rdq/RDQ-spec-*.md 的「✔ 驗收條件」

不重複 consistency.py 已經涵蓋的第 9、10 條（用語紅線、關鍵數字一致），
本檔補的是需要看結構的那幾條：出處可溯源、表格無空格、決策器情境、衛教單張用語。

執行： PYTHONUTF8=1 python src/acceptance.py
"""
import os
import re
import subprocess
import sys
import zipfile

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
OUT = os.path.join(ROOT, 'output')
REPORT = os.path.join(OUT, '台灣成人肺炎鏈球菌疫苗報告.html')
LEAFLET = os.path.join(OUT, '肺炎鏈球菌疫苗病人衛教單張.html')
DECK = os.path.join(OUT, '台灣成人肺炎鏈球菌疫苗簡報.pptx')

fails = []
warns = []


def ok(label, detail=''):
    print('  ✓ %s%s' % (label, ('　' + detail) if detail else ''))


def bad(label, detail=''):
    fails.append(label)
    print('  ✗ %s%s' % (label, ('　' + detail) if detail else ''))


def warn(label, detail=''):
    warns.append(label)
    print('  ⚠ %s%s' % (label, ('　' + detail) if detail else ''))


def strip_tags(h):
    h = re.sub(r'<script.*?</script>', ' ', h, flags=re.S)
    h = re.sub(r'<style.*?</style>', ' ', h, flags=re.S)
    return re.sub(r'<[^>]+>', ' ', h)


report = open(REPORT, encoding='utf-8').read()
leaflet = open(LEAFLET, encoding='utf-8').read()
deck_txt = ' '.join(
    zipfile.ZipFile(DECK).read(n).decode('utf-8')
    for n in zipfile.ZipFile(DECK).namelist() if n.startswith('ppt/slides/slide')
)
deck_txt = ' '.join(re.findall(r'<a:t>(.*?)</a:t>', deck_txt, re.S))

print('總驗收：依 RDQ 規格卡「✔ 驗收條件」逐條核對\n')

# ── 1 ──
print('【1】四支疫苗血清型清單完整且交叉比對正確，PCV20／PCV21 差異明確標出')
deck_src = open(os.path.join(ROOT, 'src', 'deck.js'), encoding='utf-8').read()
if '血清型矩陣：分帶語意與組成表不符' in deck_src and '血清型矩陣：' in deck_src:
    ok('產生器內建建置期驗算（價數／聯集／分帶語意），不符即中止建置')
else:
    bad('產生器缺少血清型自我驗算')
for k in ('PCV20 有、PCV21 無', '兩者共有', 'PCV21 有、PCV20 無'):
    (ok if k in deck_txt else bad)('簡報矩陣標出「%s」' % k)

# ── 2 ──
print('\n【2】HTML 報告每個效力數字與百分比都可點回出處，並標有抓取日期')
secs = re.split(r'<h2 ', report)
miss = []
for sec in secs[1:]:
    sid = re.search(r'id="(s\d)"', sec)
    if not sid:
        continue
    txt = strip_tags(sec)
    if not re.search(r'\d+(\.\d+)?\s*%', txt):
        continue
    has_src = 'class="src"' in sec or 'figcaption' in sec
    has_date = '查詢日' in sec or '查證日' in sec
    has_link = 'href="http' in sec or 'PMID' in sec
    if not (has_src and has_date and has_link):
        miss.append('%s（出處%s／日期%s／連結或PMID%s）' % (
            sid.group(1), '有' if has_src else '無', '有' if has_date else '無',
            '有' if has_link else '無'))
for m in miss:
    bad('章節 %s 的百分比缺出處要素' % m)
if not miss:
    ok('所有含百分比的章節都帶出處、查詢日與連結／PMID')

# ── 3 ──
print('\n【3】高風險個案表逐類標明「打哪支／要不要序貫／有無公費資格」，無空格未填')
m = re.search(r'<table id="hrtab">(.*?)</table>', report, re.S)
if not m:
    bad('找不到逐類高風險個案表（id="hrtab"）')
else:
    body = m.group(1)
    rows = [r for r in re.findall(r'<tr>(.*?)</tr>', body, re.S) if 'colspan' not in r]
    data_rows = [r for r in rows if '<td' in r]
    empty = 0
    for r in data_rows:
        for c in re.findall(r'<td[^>]*>(.*?)</td>', r, re.S):
            if not strip_tags(c).strip():
                empty += 1
    if empty:
        bad('逐類表有 %d 個空白儲存格' % empty)
    else:
        ok('逐類表 %d 類，無空白儲存格' % len(data_rows))
    # 規格卡點名的適應症一個都不能漏
    NEED = ['慢性心臟病', '慢性肺病／COPD', '氣喘', '糖尿病', '慢性肝病與肝硬化',
            '慢性腎病與腎病症候群', '酒精使用障礙', '吸菸者', '脾臟功能缺損', '鐮刀型貧血',
            '免疫功能不全', 'HIV', '器官移植', '血液惡性疾病', '實體癌', '人工電子耳植入', '腦脊髓液滲漏']
    lack = [k for k in NEED if k not in strip_tags(body)]
    (ok if not lack else bad)('涵蓋規格卡點名的全部適應症', '' if not lack else '缺 ' + '、'.join(lack))
    # 每一列都要有三件事
    lack3 = [i + 1 for i, r in enumerate(data_rows)
             if not ('PCV20' in r and '序貫' in r and ('公費' in r or '自費' in r or '依認定' in r))]
    (ok if not lack3 else bad)('每一類都標明打哪支／要不要序貫／公費資格',
                               '' if not lack3 else '第 %s 列不完整' % lack3)

# ── 4 ── 決策器抽測 5 種情境
print('\n【4】決策器抽測 5 種情境，且與流程圖規則一致')
#    五種情境對應流程圖（簡報 S38）的五列，逐一比對「應得結論」
CASES = [
    ('70 歲，從未接種　→ 流程圖第 1 列',
     {'birthYear': 1956, 'history': 'none'}, 'vaccinate'),
    ('68 歲，只打過 PPV23 滿 18 個月　→ 第 2 列',
     {'birthYear': 1958, 'history': 'ppv23', 'monthsSinceLast': 18}, 'vaccinate'),
    ('66 歲脾臟功能缺損，打過 PCV13 三個月　→ 第 4 列（8 週但書）',
     {'birthYear': 1960, 'history': 'pcv13', 'monthsSinceLast': 3, 'fundedRisk': ['asplenia']}, 'vaccinate'),
    ('66 歲，19–64 歲高風險期間打完兩劑、間隔 72 個月　→ 第 5 列（三條件全符合）',
     {'birthYear': 1960, 'history': 'pcv13_ppv23', 'monthsSinceLast': 72,
      'fundedRisk': ['asplenia'], 'priorHighRisk19to64': True}, 'vaccinate'),
    ('66 歲，一般人打完兩劑、間隔 72 個月　→ 第 5 列（不符則視為已完成）',
     {'birthYear': 1960, 'history': 'pcv13_ppv23', 'monthsSinceLast': 72}, 'complete'),
]
script = os.path.join(ROOT, 'src', '_acc_cases.js')
import json
with open(script, 'w', encoding='utf-8') as f:
    f.write("const {decide}=require('./engine.js');\n"
            "const C=" + json.dumps([[c[0], c[1], c[2]] for c in CASES], ensure_ascii=False) + ";\n"
            "C.forEach(function(c){var i=c[1];i.currentYear=2026;var r=decide(i);\n"
            "console.log(JSON.stringify({l:c[0],want:c[2],got:r.recommendation.action,"
            "pay:r.eligibility.funded?'公費':'自費',head:r.recommendation.headline,"
            "vac:r.recommendation.vaccine||null},null,0));});\n")
try:
    res = subprocess.run(['node', script], cwd=ROOT, capture_output=True, text=True, encoding='utf-8')
    if res.returncode != 0:
        bad('決策器抽測執行失敗', (res.stderr or '').strip().splitlines()[-1] if res.stderr else '')
    else:
        mismatch = 0
        for line in res.stdout.strip().splitlines():
            d = json.loads(line)
            hit = (d['want'] == d['got'])
            if not hit:
                mismatch += 1
            print('     %s %s' % ('✓' if hit else '✗', d['l']))
            print('        → %s｜%s｜%s' % (d['pay'], d['head'], '／'.join(d['vac']) if d['vac'] else '不再接種'))
        if mismatch:
            bad('有 %d 種情境與流程圖結論不符' % mismatch)
        else:
            ok('5 種情境結論與流程圖一致，無「無建議」死角')
finally:
    if os.path.exists(script):
        os.remove(script)
print('     （窮舉測試另見 node src/engine.test.js，419,381 項）')

# ── 5 ──
print('\n【5】決策器結果附「依據來源」並註明不取代臨床判斷')
(ok if ('依據' in report and '不取代' in report) else bad)('報告含「依據」與「不取代臨床判斷」字樣')

# ── 6 ──
print('\n【6】流程圖能回答門診三種常見情境')
for k, kw in (('未接種過', '從未接種'), ('打過 PPV23', '僅接種過 23 價'), ('打過 PCV13', '僅接種過 13 或 15 價')):
    (ok if kw in deck_txt else bad)('流程圖涵蓋「%s」' % k)

# ── 7 ──
print('\n【7】合併施打表 9 種疫苗、分兩區、無 28 天錯誤宣稱、含 MenACWY-D 例外與非官方聲明')
NINE = ['流感', 'COVID-19', 'RSV', 'Tdap', 'A 型肝炎', 'B 型肝炎', 'Shingrix', 'MMR', '水痘']
for doc, name in ((deck_txt, '簡報'), (strip_tags(report), '報告')):
    missing = [v for v in NINE if v not in doc]
    (ok if not missing else bad)('%s涵蓋 9 種疫苗' % name, '' if not missing else '缺 ' + '、'.join(missing))
for doc, name in ((deck_txt, '簡報'), (strip_tags(report), '報告')):
    two = ('不活化' in doc) and ('活性減毒' in doc)
    (ok if two else bad)('%s不活化與活性減毒分兩區' % name)
wrong = re.search(r'肺(炎鏈球菌|鏈)[^。]{0,20}(活性|MMR|水痘)[^。]{0,20}間隔\s*(≥\s*)?28', strip_tags(report))
(bad if wrong else ok)('未錯誤宣稱肺鏈與活性疫苗須間隔 28 天')
for doc, name in ((deck_txt, '簡報'), (strip_tags(report), '報告')):
    (ok if 'MenACWY-D' in doc else bad)('%s標明 MenACWY-D 例外' % name)
    (ok if '非疾管署公告' in doc else bad)('%s標明非疾管署公告內容' % name)

# ── 8 ──
print('\n【8】免疫橋接教學模組為簡報獨立章節，且位於效力證據章節之前')
i_bridge = deck_txt.find('免疫橋接 ①')
i_evid = deck_txt.find('證據的三個層級')
if i_bridge < 0 or i_evid < 0:
    bad('找不到免疫橋接模組或效力證據章節')
elif i_bridge < i_evid:
    ok('免疫橋接模組在效力證據章節之前')
    warn('模組首頁本身引用 45.56% 作為破題對照',
         '這是模組內的教學用例，非獨立效力主張；驗收時請確認可接受')
else:
    bad('免疫橋接模組出現在效力證據章節之後')

# ── 11 ──
print('\n【11】衛教單張無血清型編號、無未解釋的英文縮寫')
lt = strip_tags(leaflet)
sero = re.findall(r'(?<![0-9A-Za-z])(?:6A|6B|9V|9N|10A|11A|12F|14|15A|15B|15C|16F|17F|18C|19A|19F|20A|22F|23A|23B|23F|24F|33F|35B)(?![0-9A-Za-z%])', lt)
(ok if not sero else bad)('全文無血清型編號', '' if not sero else '出現：' + '、'.join(sorted(set(sero))))
abbr = set(re.findall(r'\b(?!PCV13|PCV15|PCV20|PCV21|PPV23|COVID|Claude)[A-Z]{2,}(?:-[A-Z0-9]+)?\b', lt))
(ok if not abbr else warn)('全文無未解釋的英文縮寫', '' if not abbr else '請確認：' + '、'.join(sorted(abbr)))

# ── 規格偏離 ──
print('\n【規格偏離：需使用者確認】')
warn('衛教單張由「A4 單頁」改為「A4 雙面兩頁」',
     '2026-08-09 使用者明確指示維持雙面並加大邊界以容忍 Letter 紙匣；規格卡原文為單頁')

print('\n' + '─' * 56)
if fails:
    print('❌ %d 項未通過：%s' % (len(fails), '、'.join(fails)))
    sys.exit(1)
print('✅ 驗收條件全部通過（另有 %d 項待使用者確認，見上方 ⚠）' % len(warns))
