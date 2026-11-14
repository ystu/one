"""Markdown 三層大綱轉 Word。

以既有大綱 Word 為格式參考（沿用其樣式、多層次編號與字型設定），
依 docs/模板與提示詞/大綱整理規則.md 第十一節產生標題、署名與三層編號段落。
Markdown 中的引用區塊（> 開頭的討論紀錄）不會輸出。

用法：
    python3 scripts/md_outline_to_docx.py 來源大綱.md 輸出大綱.docx
    python3 scripts/md_outline_to_docx.py 來源大綱.md 輸出大綱.docx --narrow --line 21.5

選項：
    --reference  格式參考 Word，預設為三曹普渡大綱v1（已含三層多層次編號）
    --narrow     上下左右 1.27 公分窄邊界（一張 A4 版用）；預設為一般版邊界
    --line       內文固定行高（點），預設 21.5
    --levels     輸出層數（2 或 3），預設 3；2 只輸出大題與次標題
    --text-numbering  標號以文字輸出，不使用 Word 自動編號；Google Drive 等
                 不支援中文編號格式的檢視器會把「一、（一）」顯示成「1、（1）」，
                 需要在這類環境開啟時使用
"""
import argparse
import html
import re
import zipfile
from pathlib import Path

REPO = Path(__file__).resolve().parent.parent
DEFAULT_REF = REPO / 'speech/three-realms-salvation/三曹普渡大綱v1_凃信埕.docx'
AUTHOR = '後學 凃信埕 敬述'

# 各層：左縮排、凸排、字元間距（twips / 1/20 點），沿用參考檔設定
LEVELS = {0: (720, 720, 0), 1: (1440, 960, 20), 2: (1680, 480, 20)}
MARGINS = {'normal': (1440, 1800), 'narrow': (720, 720)}  # (上下, 左右)


def parse_outline(md_path):
    title, items = None, []
    for raw in Path(md_path).read_text(encoding='utf8').splitlines():
        s = raw.strip()
        if not s or s.startswith('>'):
            continue
        s = s.lstrip('#').strip()
        if s.startswith('主題：'):
            title = s[3:].strip()
        elif m := re.match(r'^[一二三四五六七八九十]+、\s*(.+)$', s):
            items.append((0, m.group(1)))
        elif m := re.match(r'^（[一二三四五六七八九十]+）\s*(.+)$', s):
            items.append((1, m.group(1)))
        elif m := re.match(r'^\d+\.\s+(.+)$', s):
            items.append((2, m.group(1)))
        else:
            raise SystemExit(f'無法辨識的行（大綱只允許三層標號）：{s}')
    if not title or not items:
        raise SystemExit('找不到「主題：」或大綱內容')
    return title, items


def run(text, size, bold=False, spacing=20):
    b = '<w:b /><w:bCs />' if bold else '<w:b w:val="0" /><w:bCs w:val="0" />'
    return ('<w:r><w:rPr><w:rFonts w:ascii="標楷體" w:hAnsi="標楷體" w:eastAsia="標楷體" w:cs="標楷體" w:hint="eastAsia" />'
            f'{b}<w:color w:val="000000" /><w:spacing w:val="{spacing}" /><w:sz w:val="{size}" /><w:szCs w:val="{size}" />'
            f'</w:rPr><w:t xml:space="preserve">{html.escape(text, quote=False)}</w:t></w:r>')


CN = '一二三四五六七八九十'


def cn_num(n):
    if n <= 10:
        return CN[n - 1]
    if n < 20:
        return '十' + CN[n - 11]
    return CN[n // 10 - 1] + '十' + (CN[n % 10 - 1] if n % 10 else '')


def label(lv, n):
    return [f'{cn_num(n)}、', f'（{cn_num(n)}）', f'{n}.'][lv]


def build(title, items, ref, out, narrow, line_pt, levels=3, text_numbering=False):
    with zipfile.ZipFile(ref) as z:
        doc = z.read('word/document.xml').decode('utf8')
    numid = re.search(r'<w:numId w:val="(\d+)"', doc).group(1)
    line = round(line_pt * 20)

    paras = [
        '<w:p><w:pPr><w:jc w:val="center" /><w:spacing w:before="0" w:after="0" w:lineRule="exact" w:line="560" />'
        f'<w:snapToGrid w:val="0" /></w:pPr>{run(title, 40, bold=True)}</w:p>',
        '<w:p><w:pPr><w:jc w:val="right" /><w:spacing w:before="0" w:after="0" w:lineRule="exact" w:line="360" />'
        f'<w:snapToGrid w:val="0" /></w:pPr>{run(AUTHOR, 24)}</w:p>',
    ]
    seen_first = False
    items = [(lv, text) for lv, text in items if lv < levels]
    counters = [0, 0, 0]
    for lv, text in items:
        left, hang, sp = LEVELS[lv]
        counters[lv] += 1
        counters[lv + 1:] = [0] * (2 - lv)
        before = 60 if lv == 0 and seen_first else 0  # 第二個大題起段前 3 點
        seen_first = seen_first or lv == 0
        keep = '<w:keepNext />' if lv < levels - 1 else ''
        if text_numbering:
            num = ''
            tabs = f'<w:tabs><w:tab w:val="left" w:pos="{left}" /></w:tabs>'
            body = run(label(lv, counters[lv]), 28, spacing=sp) + '<w:r><w:tab /></w:r>' + run(text, 28, spacing=sp)
        else:
            num = f'<w:numPr><w:ilvl w:val="{lv}" /><w:numId w:val="{numid}" /></w:numPr>'
            tabs = f'<w:tabs><w:tab w:val="num" w:pos="{left}" /></w:tabs>'
            body = run(text, 28, spacing=sp)
        paras.append(
            f'<w:p><w:pPr>{keep}{num}'
            f'<w:spacing w:before="{before}" w:after="0" w:lineRule="exact" w:line="{line}" />'
            f'{tabs}<w:ind w:left="{left}" w:hanging="{hang}" />'
            f'<w:snapToGrid w:val="0" /></w:pPr>{body}</w:p>')

    tb, lr = MARGINS['narrow' if narrow else 'normal']
    sect = re.search(r'<w:sectPr.*?</w:sectPr>', doc).group(0)
    sect = re.sub(r'<w:pgMar [^>]*/>',
                  f'<w:pgMar w:top="{tb}" w:right="{lr}" w:bottom="{tb}" w:left="{lr}" '
                  'w:header="851" w:footer="992" w:gutter="0" />', sect)
    doc = re.sub(r'<w:body>.*</w:body>', '<w:body>' + ''.join(paras) + sect + '</w:body>', doc, flags=re.S)

    with zipfile.ZipFile(ref) as zin, zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            if info.filename == 'word/document.xml':
                data = doc.encode('utf8')
            elif info.filename == 'docProps/core.xml':
                data = re.sub(r'<dc:title>.*?</dc:title>', f'<dc:title>{html.escape(title)}</dc:title>',
                              data.decode('utf8')).encode('utf8')
            zout.writestr(info, data)


def main():
    ap = argparse.ArgumentParser(description='Markdown 三層大綱轉 Word')
    ap.add_argument('source')
    ap.add_argument('output')
    ap.add_argument('--reference', default=str(DEFAULT_REF))
    ap.add_argument('--narrow', action='store_true')
    ap.add_argument('--line', type=float, default=21.5)
    ap.add_argument('--levels', type=int, choices=(2, 3), default=3)
    ap.add_argument('--text-numbering', action='store_true')
    a = ap.parse_args()
    title, items = parse_outline(a.source)
    build(title, items, a.reference, a.output, a.narrow, a.line, a.levels, a.text_numbering)
    counts = [sum(1 for lv, _ in items if lv == i) for i in range(a.levels)]
    labels = ['大題', '次標題', '重點'][:a.levels]
    print(f'{title}：' + '、'.join(f'{l} {c}' for l, c in zip(labels, counts)) + f' → {a.output}')


if __name__ == '__main__':
    main()
