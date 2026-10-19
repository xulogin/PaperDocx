# -*- coding: utf-8 -*-
r"""把 docx 导出的 PDF 里**真正印出来的**参考文献表，和稿件 md 里作者写的 65 条逐条比。

    py -3.12 -X utf8 排版/比文献表.py 排版/_预览/_预览_V49.pdf

★ 「跑通」≠「对」。出 Word 不报错只说明字段拼对了，不说明印出来的文献和作者
  写的是同一条。Zotero 条目的刊名缩写、条目类型（preprint / report）、
  页码范围、出版社，都会让印出来的那一行和 md 里不一样。
"""
from __future__ import annotations

import difflib
import re
import sys
from pathlib import Path

import fitz

这里 = Path(__file__).resolve().parent
sys.path.insert(0, str(这里))
import _项目                                            # noqa: E402

根 = _项目.找根()
pdf = Path(sys.argv[1] if len(sys.argv) > 1
           else 这里 / "_预览" / next(p.name for p in (这里 / "_预览").glob("*.pdf")))
if not pdf.is_absolute():
    pdf = 根 / pdf

md = _项目.三件套(根)[2].read_text("utf-8")
稿 = [l.strip() for l in md.partition("\n## References\n")[2].split("\n") if l.strip()]

d = fitz.open(pdf)
文 = "\n".join(p.get_text() for p in d)
始 = 文.rindex("\nReferences")
尾 = 文[始:]
# 去页码行和页眉，再按「姓, 名缩写., ... 年.」开头重新断条
尾 = re.sub(r"\n\s*\d{1,3}\s*\n", "\n", 尾)
行 = " ".join(尾.split("\n")[1:])
行 = re.sub(r"\s+", " ", 行)
起 = [m.start() for m in re.finditer(
    # ★ 作者多的条目（Puliti 那条有 33 个作者）光作者串就 400 多字符，
    #   窗口开小了会把下一条并进上一条，看着像「排版把两条粘一起了」。
    r"(?<![A-Za-z,])[A-ZÅÄÖÉÜ][\wÀ-ɏ'’\-]+(?:,\s*[A-Z]\.(?:-?[A-Z]\.)*){1,4},?.{0,900}?\b(?:19|20)\d{2}[a-z]?\.\s", 行)]
印 = []
for i, s in enumerate(起):
    印.append(行[s:起[i + 1] if i + 1 < len(起) else len(行)].strip())

print(f"md 里 {len(稿)} 条；PDF 里断出 {len(印)} 条")


def 规(s, 保大小写=False):
    r"""★ 默认压成小写是为了配对；配上之后**必须再按原样比一次大小写** ——
    实测教训：Zotero 里 Bonan 2008 的刊名存成小写 `science`，压成小写之后
    和稿子里的 `Science` 一模一样，报「逐字相同」，可印出来就是小写。"""
    s = s.replace("‐", "-").replace("–", "-").replace("—", "-").replace("’", "'")
    if not 保大小写:
        s = s.lower()
    return re.sub(r"\s+", " ", re.sub(r"[^A-Za-z0-9]+", " ", s)).strip()


印规 = [规(x) for x in 印]
用过 = set()
差, 大小写 = [], []
for s in 稿:
    g = 规(s)
    最好, 最分 = None, 0.0
    for j, p in enumerate(印规):
        if j in 用过:
            continue
        r = difflib.SequenceMatcher(None, g[:150], p[:150]).ratio()
        if r > 最分:
            最好, 最分 = j, r
    if 最好 is None:
        差.append((s, "", 0.0))
        continue
    用过.add(最好)
    if 最分 < 0.995:
        差.append((s, 印[最好], 最分))
    elif 规(s, True) != 规(印[最好], True):
        大小写.append((s, 印[最好]))
没配上 = [印[j] for j in range(len(印)) if j not in 用过]

print(f"逐字相同 {len(稿) - len(差)} 条；有出入 {len(差)} 条；PDF 里多出 {len(没配上)} 条\n")
for s, p, r in sorted(差, key=lambda x: x[2]):
    print(f"[{r:.3f}]\n  md ：{s}\n  印出：{p}")
    sm = difflib.SequenceMatcher(None, s, p)
    细 = [f"{t}: {s[i1:i2]!r} → {p[j1:j2]!r}"
         for t, i1, i2, j1, j2 in sm.get_opcodes() if t != "equal"]
    print("  差：" + " | ".join(细[:8]) + "\n")
for p in 没配上:
    print("PDF 里多出：", p[:140])
if 大小写:
    print(f"\n★ 只差大小写的 {len(大小写)} 条（Zotero 字段里存的就是这个大小写）：")
    for s, p in 大小写:
        a, b = 规(s, True).split(), 规(p, True).split()
        不同 = [f"{x} → {y}" for x, y in zip(a, b) if x != y]
        print(f"  {s[:60]}…\n     {' | '.join(不同[:6])}")
