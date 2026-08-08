# -*- coding: utf-8 -*-
"""
跨產出一致性檢查

對應規格卡驗收條件：
  「三份產出（HTML／PPTX／衛教單張）的關鍵數字完全一致，無互相矛盾」
  「三份產出全文逐句掃過用語紅線」

執行： PYTHONUTF8=1 python src/consistency.py
"""
import os
import re
import sys
import zipfile

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
OUT = os.path.join(ROOT, 'output')

# ── 關鍵數字：每一項都必須在「有提到該主題」的產出中一致 ──
KEY_NUMBERS = {
    'CAPiTA 收案人數': '84,496',
    'CAPiTA 主要終點 VE': '45.56%',
    'CAPiTA VT-IPD VE': '75.00%',
    'CAPiTA 信賴區間': '95.2%',
    'STRIDE-3 隨機人數': '2,663',
    'IPD 發生率 2024': '1.25',
    '完整接種比率': '12%',
    'PCV20 獨有型數': '9',
    'PCV21 獨有型數': '11',
    '共有型數': '10',
    '活性疫苗間隔': '28',
    '新制上路日': '8 月 10 日',
}

# ── 用語紅線（data/07 第 ⑤ 節）──
RED_LINES = [
    r'PCV21[^。]{0,12}保護力較好',
    r'PCV21[^。]{0,12}較能預防肺炎',
    r'PCV2[01][^。]{0,10}效力\s*\d+\s*%',
    r'保護力提升\s*\d+\s*倍',
    r'PCV13[^。]{0,20}降低肺炎住院',
]

# ── 必須出現的但書（避免誤導）──
REQUIRED_CAVEATS = {
    'HTML': ['免疫橋接', '探索性終點', '非疾管署公告'],
    'PPTX': ['免疫橋接', '探索性終點', '非疾管署公告'],
}


def read_html(p):
    with open(p, encoding='utf-8') as f:
        s = f.read()
    s = re.sub(r'<script.*?</script>', ' ', s, flags=re.S)
    s = re.sub(r'<style.*?</style>', ' ', s, flags=re.S)
    return re.sub(r'<[^>]+>', ' ', s)


def read_pptx(p):
    out = []
    with zipfile.ZipFile(p) as z:
        for n in sorted(x for x in z.namelist() if re.match(r'ppt/(slides|notesSlides)/\w+\d+\.xml$', x)):
            xml = z.read(n).decode('utf-8', 'ignore')
            out.extend(re.findall(r'<a:t>(.*?)</a:t>', xml, flags=re.S))
    return ' '.join(out)


def main():
    docs = {}
    html = os.path.join(OUT, '台灣成人肺炎鏈球菌疫苗報告.html')
    pptx = os.path.join(OUT, '台灣成人肺炎鏈球菌疫苗簡報.pptx')
    leaflet = os.path.join(OUT, '肺炎鏈球菌疫苗病人衛教單張.html')

    if os.path.exists(html):
        docs['HTML'] = read_html(html)
    if os.path.exists(pptx):
        docs['PPTX'] = read_pptx(pptx)
    if os.path.exists(leaflet):
        docs['衛教單張'] = read_html(leaflet)

    if not docs:
        print('✗ 找不到任何產出檔')
        return 1

    print('檢查產出：' + '、'.join(docs.keys()) + '\n')
    fails = 0

    # 1. 關鍵數字一致性（僅比對「主要技術文件」HTML 與 PPTX；衛教單張刻意不列技術數字）
    print('【1】關鍵數字一致性（HTML vs PPTX）')
    tech = {k: v for k, v in docs.items() if k in ('HTML', 'PPTX')}
    for label, num in KEY_NUMBERS.items():
        present = {k: (num in v) for k, v in tech.items()}
        if len(set(present.values())) > 1:
            missing = [k for k, ok in present.items() if not ok]
            print('  ⚠ %s（%s）：%s 未出現' % (label, num, '、'.join(missing)))
        elif all(present.values()):
            print('  ✓ %s = %s' % (label, num))
    print()

    # 2. 用語紅線
    print('【2】用語紅線掃描')
    for name, text in docs.items():
        hit = []
        for pat in RED_LINES:
            for m in re.finditer(pat, text):
                # 「不可以說」清單本身會列出這些句子，屬於教學用途，需排除
                ctx = text[max(0, m.start() - 40):m.start()]
                if '不可以說' in ctx or '不可寫' in ctx or '⛔' in ctx or '不得出現' in ctx or '不可' in ctx[-12:]:
                    continue
                hit.append(m.group(0)[:40])
        if hit:
            fails += 1
            print('  ✗ %s 出現紅線用語：%s' % (name, '｜'.join(hit)))
        else:
            print('  ✓ %s 無違規用語' % name)
    print()

    # 3. 必要但書
    print('【3】必要但書')
    for name, needed in REQUIRED_CAVEATS.items():
        if name not in docs:
            continue
        for c in needed:
            if c in docs[name]:
                print('  ✓ %s 含「%s」' % (name, c))
            else:
                fails += 1
                print('  ✗ %s 缺少「%s」' % (name, c))
    print()

    # 4. 血清型組成與疫苗語意色：簡報產生器與 HTML 報告必須逐字同源
    #    （曾發生報告把血清型 20 與 20A 拆成兩欄、與自己的內文矛盾，也與簡報不一致）
    print('【4】血清型組成與疫苗語意色（deck.js vs report.template.html）')
    deck_src = open(os.path.join(ROOT, 'src', 'deck.js'), encoding='utf-8').read()
    rep_src = open(os.path.join(ROOT, 'src', 'report.template.html'), encoding='utf-8').read()

    def sero_sets(text):
        got = {}
        for v in ('PPV23', 'PCV13', 'PCV20', 'PCV21'):
            m = re.search(v + r"\s*:\s*\[([^\]]*)\]", text)
            if m:
                got[v] = set(re.findall(r"'([^']+)'", m.group(1)))
        return got

    a, b = sero_sets(deck_src), sero_sets(rep_src)
    valency = {'PPV23': 23, 'PCV13': 13, 'PCV20': 20, 'PCV21': 21}
    for v, n in valency.items():
        if v not in a or v not in b:
            fails += 1
            print('  ✗ %s 在其中一份找不到組成定義' % v)
        elif a[v] != b[v]:
            fails += 1
            print('  ✗ %s 兩份不一致，差異：%s' % (v, sorted(a[v] ^ b[v])))
        elif len(a[v]) != n:
            fails += 1
            print('  ✗ %s 型數為 %d，應為 %d' % (v, len(a[v]), n))
        else:
            print('  ✓ %s 兩份一致且為 %d 型' % (v, n))

    for key, var in (('PPV23', 'ppv23'), ('PCV13', 'pcv13'), ('PCV20', 'pcv20'), ('PCV21', 'pcv21')):
        md = re.search(key + r"\s*:\s*'([0-9A-Fa-f]{6})'", deck_src)
        mr = re.search(r"--" + var + r"\s*:\s*#([0-9A-Fa-f]{6})", rep_src)
        if not md or not mr:
            fails += 1
            print('  ✗ %s 語意色在其中一份找不到' % key)
        elif md.group(1).lower() != mr.group(1).lower():
            fails += 1
            print('  ✗ %s 語意色不一致：deck #%s vs report #%s' % (key, md.group(1), mr.group(1)))
        else:
            print('  ✓ %s 語意色一致 #%s' % (key, md.group(1).lower()))
    print()

    print('─' * 52)
    if fails:
        print('❌ %d 項未通過' % fails)
        return 1
    print('✅ 一致性檢查通過')
    return 0


if __name__ == '__main__':
    sys.exit(main())
