# -*- coding: utf-8 -*-
r"""排版后逐项核对：md 原稿里的东西，docx 里一样不能少。

    py -3.12 -X utf8 排版/核原文.py V63_英文版.md 交付/V63_排版版.docx [--图目录]

核：段落和表格的每段文字、全部数字（次数也要对）、公式、行内公式、图（加 --图目录 时逐字节比图文件，
只对原图版有意义）。md 的参考文献表不比（文献表由 Zotero 生成，用 比文献表.py）。缺东西时退出码 1。
"""
import hashlib
import html
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path

标记 = r"[*^~⟦⟧\\]"
行内式 = r"(?<!\$)\$(?!\$)(?:[^$\n]|\\\$)+?\$(?!\$)"


def 规(s):
    s = re.sub(r"[\s ​]+", "", html.unescape(s))
    return re.sub(标记, "", s)


def 数字(s):
    return Counter(re.findall(r"\d+(?:\.\d+)?", s))


def 读md(路径):
    s = Path(路径).read_text("utf-8").replace("\r\n", "\n")
    s = re.split(r"^#+\s*References\s*$", s, flags=re.M)[0]
    s = re.sub(r"<!--.*?-->", "", s, flags=re.S)
    公式 = len(re.findall(r"^\$\$\s*$", s, flags=re.M)) // 2
    s = re.sub(r"^\$\$\s*\n.*?\n\$\$\s*$", "", s, flags=re.S | re.M)
    图 = re.findall(r"!\[[^\]]*\]\(([^)]+)\)", s)
    s = re.sub(r"!\[[^\]]*\]\([^)]+\)(\{[^}]*\})?", "", s)
    行内 = len(re.findall(行内式, s))
    无式 = re.sub(行内式, "\n\n", s)
    块 = []
    for b in re.split(r"\n\s*\n", 无式):
        for 行 in b.split("\n"):
            行 = 行.strip()
            if not 行 or re.fullmatch(r"[|:\- ]+", 行):
                continue
            行 = re.sub(r"^#+\s*|^\*\*(?:Authors|Affiliations|Correspondence):\*\*\s*", "", 行)
            if not 行:
                continue
            if 行.startswith("|"):
                块 += [c.strip() for c in 行.strip("|").split("|") if c.strip()]
            else:
                块.append(行)
    return 块, 无式, 公式, 行内, 图


def 读docx(路径):
    z = zipfile.ZipFile(路径)
    x = z.read("word/document.xml").decode("utf-8")
    x = re.sub(r"<w:instrText[^>]*>.*?</w:instrText>", "", x, flags=re.S)
    公式段 = len(re.findall(r"<m:oMathPara[ >]", x))
    全式 = len(re.findall(r"<m:oMath[ >]", x))
    无式 = re.sub(r"<m:oMath[ >].*?</m:oMath>", "", x, flags=re.S)
    段 = [html.unescape("".join(re.findall(r"<w:t(?: [^>]*)?>([^<]*)</w:t>", p)))
         for p in re.findall(r"<w:p[ >].*?</w:p>", 无式, flags=re.S)]
    if "References" in 段:
        段 = 段[: 段.index("References")]
    媒 = {hashlib.sha1(z.read(n)).hexdigest() for n in z.namelist() if n.startswith("word/media/")}
    return "\n".join(段), 公式段, 全式, 媒, len(re.findall(r"<pic:pic[ >]", x))


def main():
    md, docx = sys.argv[1], sys.argv[2]
    块, 无式, 公式, 行内, 图 = 读md(md)
    文, 公式段, 全式, 媒, 图数 = 读docx(docx)
    D = 规(文)
    缺 = [b for b in 块 if len(规(b)) >= 4 and 规(b) not in D]
    print(f"文字块 {len(块)} 个，缺 {len(缺)} 个")
    for b in 缺[:30]:
        print("  缺：", b[:110])
    nm, nd = 数字(re.sub(标记, "", 无式)), 数字(文)
    少 = {k: nm[k] - nd[k] for k in nm if nm[k] > nd[k]}
    print(f"数字 {sum(nm.values())} 个，docx 里少的：{少 or '无'}")
    print(f"公式 md {公式} / docx {公式段}；行内公式 md {行内} / docx {全式 - 公式段}；图 md {len(图)} / docx {图数}")
    差 = len(缺) + len(少) + (公式 != 公式段) + (行内 != 全式 - 公式段) + (len(图) != 图数)
    if "--图目录" in sys.argv:
        根 = Path(md).resolve().parent
        未 = [p for p in 图 if hashlib.sha1((根 / p).read_bytes()).hexdigest() not in 媒]
        print(f"图文件逐字节比：{len(图) - len(未)}/{len(图)} 相同", 未[:5])
        差 += len(未)
    print("核对通过" if not 差 else f"⛔ {差} 处要看")
    sys.exit(1 if 差 else 0)


main()
