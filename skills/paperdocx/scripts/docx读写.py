# -*- coding: utf-8 -*-
r"""在排版好的 docx 上**只读文字 / 只写文字**，样式一个都不碰。

`三方同步.py` 用它。两个方向：

    读块(doc) -> {块序: 同步串}      把 Word 段落化成能和 md 比的中间串
    写块(doc, 序, 新同步串, ...)      把新文字灌回去，段落样式（pPr）原样保留

## 怎么认出「这段是 md 的哪一块」

出Word.py 在每块起始处埋了书签 `PDX0001`（前缀见 排版配置.锚点前缀）。读的时候按书签名定位。
书签在 Word 里编辑正文不会掉；**整段删掉才会掉**，掉了就报出来让人看，不猜。

## 写回时的原则

* `w:pPr` 整个保留 —— 作者在 Word 里调的样式、缩进、间距全不动。
* run 级格式按 md 的 `**粗** *斜* ^上标^ 重建，底版 rPr 抄这段原来第一个
  纯文本 run（作者改过字体字号的话也跟着保留）。
* **行内公式（OMML）和 Zotero 引用字段原封不动搬过去**，不重建。
  所以在 Word 里改公式或改引用显示文本，同步会发现对不上并报错 —— 这是故意的，
  OMML 反推 LaTeX 不可靠，宁可报错也不许猜。
"""
from __future__ import annotations

import copy
import re
import sys
from pathlib import Path

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

这里 = Path(__file__).resolve().parent
sys.path.insert(0, str(这里))
import 块模型 as M                      # noqa: E402
import 排版配置 as C                     # noqa: E402
import _zotero字段 as Z                 # noqa: E402

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
MATH = "{http://schemas.openxmlformats.org/officeDocument/2006/math}"


# ══════════════════════════════════════════════════ 定位
def 书签表(doc):
    """{书签名: 元素}。元素是书签所在的 w:p，或书签紧跟着的 w:tbl。"""
    出 = {}
    for bm in doc.element.body.iter(qn("w:bookmarkStart")):
        名 = bm.get(qn("w:name"))
        if not re.fullmatch(r"%s\d{4}" % re.escape(C.锚点前缀), 名 or ""):
            continue
        父 = bm.getparent()
        if 父.tag == qn("w:p"):
            出[名] = 父
        else:                                   # 表前面那种：书签是 body 的直接子元素
            下 = bm.getnext()
            while 下 is not None and 下.tag not in (qn("w:p"), qn("w:tbl")):
                下 = 下.getnext()
            if 下 is not None:
                出[名] = 下
    return 出


def 段落跨度(doc, 锚元素, 条数):
    """从锚元素起连取 `条数` 个 w:p（`affil` 那种一块对多段的情况）。"""
    出, 节 = [], 锚元素
    while 节 is not None and len(出) < 条数:
        if 节.tag == qn("w:p"):
            出.append(节)
        节 = 节.getnext()
    return 出


# ══════════════════════════════════════════════════ 读
def _格式(r):
    rpr = r.find(W + "rPr")
    if rpr is None:
        return (False, False, False)
    关 = lambda e: e is not None and e.get(qn("w:val")) not in ("0", "false")
    va = rpr.find(W + "vertAlign")
    return (关(rpr.find(W + "b")),
            关(rpr.find(W + "i")),
            va is not None and va.get(qn("w:val")) == "superscript")


def 段转同步串(p, 条目id表=None):
    r"""一个 w:p → 同步串。

    `条目id表` 是 {Zotero itemKey: refN}；给了才把引用字段收成 ⟦c:refN⟧。
    """
    字段run = {}
    for f in Z.iter_fields(_包(p)):
        for i, r in enumerate(f.runs):
            字段run[id(r)] = (f, i)

    片 = []           # [(格式 或 None, 文本)]；格式 None 表示是哨兵，原样放
    数学 = 0
    for 子 in p:
        if 子.tag == qn("w:r") and id(子) in 字段run:
            f, i = 字段run[id(子)]
            if i:                               # 字段的后续 run 跳过
                continue
            if 条目id表 is None:
                片.append((None, "⟦c:?⟧"))
            else:
                键 = f.item_keys()
                refs = [条目id表.get(k, "?" + k) for k in 键]
                片.append((None, "⟦c:%s⟧" % ";".join(refs)))
            continue
        if 子.tag == qn("w:r"):
            if 子.find(MATH + "oMath") is not None:
                数学 += 1
                片.append((None, "⟦m%d⟧" % 数学))
                continue
            t = "".join(e.text or "" for e in 子.findall(W + "t"))
            if 子.find(W + "tab") is not None:
                t += "\t"
            if 子.find(W + "br") is not None:
                t += "\n"
            if t:
                片.append((_格式(子), t))
        elif 子.tag == MATH + "oMath":
            数学 += 1
            片.append((None, "⟦m%d⟧" % 数学))
        elif 子.tag == qn("w:hyperlink"):
            for r in 子.findall(W + "r"):
                t = "".join(e.text or "" for e in r.findall(W + "t"))
                if t:
                    片.append((_格式(r), t))

    # Word 会把一句话无缘无故切成好几个 run，先把同格式的相邻 run 并起来再加哨兵
    out, 上 = [], None
    缓 = ""
    def 刷():
        nonlocal 缓
        if not 缓:
            return
        b, i, s = 上
        文 = 缓
        if s:
            文 = M.上开 + 文 + M.上关
        if i:
            文 = M.斜开 + 文 + M.斜关
        if b:
            文 = M.粗开 + 文 + M.粗关
        out.append(文)
        缓 = ""

    for 格, 文 in 片:
        if 格 is None:
            刷()
            out.append(文)
            上 = None
            continue
        if 格 != 上:
            刷()
            上 = 格
        缓 += 文
    刷()
    return "".join(out)


class _包:
    """iter_fields 要的是 python-docx 的 Paragraph；这里只用到 ._p。"""
    def __init__(self, p):
        self._p = p


def 表转格(tbl, 条目id表=None):
    """w:tbl → [[同步串]]（按行列）。

    ★ `条目id表` 必须往下传 —— 表格单元格里也有 Zotero 引用（表 10 每行都有），
      不传的话读出来全是 `⟦c:?⟧`，和 md 一比就成了假冲突。
    """
    出 = []
    for tr in tbl.findall(W + "tr"):
        行 = []
        for tc in tr.findall(W + "tc"):
            行.append(" ".join(段转同步串(p, 条目id表)
                              for p in tc.findall(W + "p")).strip())
        出.append(行)
    return 出


# ══════════════════════════════════════════════════ 写
# w:rPr 里子元素的顺序是 schema 规定死的，顺序错了 Word 开门就说「文件已损坏」。
# ★ 踩过：`写段` 里先 deepcopy 底版 rPr（含 rFonts/sz/szCs）再 append b/i，
#   结果成了 rFonts,sz,szCs,b,i —— b 必须在 sz 前面。
_RPR序 = {qn("w:" + n): i for i, n in enumerate(
    "rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint "
    "noProof snapToGrid vanish webHidden color spacing w kern position sz szCs highlight u "
    "effect bdr shd fitText vertAlign rtl cs em lang eastAsianLayout specVanish oMath".split())}


def 排rPr序(rpr):
    rpr[:] = sorted(rpr, key=lambda c: _RPR序.get(c.tag, 999))


def _底版rpr(p):
    """抄这段第一个带文字的普通 run 的字符格式，当新 run 的底版。"""
    for r in p.findall(W + "r"):
        if r.find(W + "fldChar") is not None or r.find(W + "instrText") is not None:
            continue
        if r.find(MATH + "oMath") is not None:
            continue
        if "".join(e.text or "" for e in r.findall(W + "t")):
            rpr = r.find(W + "rPr")
            return copy.deepcopy(rpr) if rpr is not None else None
    return None


def 写段(p, 新串, 条目id表=None):
    r"""用新同步串重建这一段的 run 序列。`w:pPr` 原样保留。

    原来的 OMML 和 Zotero 字段按出现顺序复用；新串里引用/公式的**个数和顺序**
    必须和原来一致，不一致就抛错，不许猜。
    """
    旧数学, 旧字段 = [], []
    字段run = {}
    for f in Z.iter_fields(_包(p)):
        旧字段.append([copy.deepcopy(r) for r in f.runs])
        for r in f.runs:
            字段run[id(r)] = True
    for 子 in p:
        if 子.tag == MATH + "oMath":
            旧数学.append(copy.deepcopy(子))
        elif 子.tag == qn("w:r") and id(子) not in 字段run:
            om = 子.find(MATH + "oMath")
            if om is not None:
                旧数学.append(copy.deepcopy(om))

    底 = _底版rpr(p)
    ppr = p.find(W + "pPr")
    # ★ 书签必须留着！清段落内容时连 bookmarkStart/End 一起删掉的话，
    #   下一轮同步就认不回这一块了（实测报 KeyError: 'PDX0018'）。
    书签 = [子 for 子 in p
          if 子.tag in (qn("w:bookmarkStart"), qn("w:bookmarkEnd"))]
    for 子 in list(p):
        if 子.tag != qn("w:pPr"):
            p.remove(子)
    开 = [b for b in 书签 if b.tag == qn("w:bookmarkStart")]
    关 = [b for b in 书签 if b.tag == qn("w:bookmarkEnd")]
    for b in 开:
        p.append(b)

    def 新run(文, 粗=False, 斜=False, 上=False, 红=False):
        r = OxmlElement("w:r")
        rpr = copy.deepcopy(底) if 底 is not None else OxmlElement("w:rPr")
        for 名, 开 in (("w:b", 粗), ("w:i", 斜)):
            for old in list(rpr.findall(qn(名))):
                rpr.remove(old)
            if 开:
                rpr.append(OxmlElement(名))
        for old in list(rpr.findall(qn("w:vertAlign"))):
            rpr.remove(old)
        if 上:
            v = OxmlElement("w:vertAlign")
            v.set(qn("w:val"), "superscript")
            rpr.append(v)
        for old in list(rpr.findall(qn("w:color"))):
            rpr.remove(old)
        if 红:
            c = OxmlElement("w:color")
            c.set(qn("w:val"), "C00000")        # Word 自带的「深红」，打印也看得清
            rpr.append(c)
        排rPr序(rpr)
        if len(rpr):
            r.append(rpr)
        t = OxmlElement("w:t")
        t.set(qn("xml:space"), "preserve")
        t.text = 文
        r.append(t)
        p.append(r)

    用数学, 用字段 = 0, 0
    i, n = 0, len(新串)
    粗 = 斜 = 上 = 红 = False
    缓 = ""
    while i < n:
        m = M.哨兵.match(新串, i)
        if not m:
            j = M.哨兵.search(新串, i)
            末 = j.start() if j else n
            缓 += 新串[i:末]
            i = 末
            continue
        if 缓:
            新run(缓, 粗, 斜, 上, 红)
            缓 = ""
        t, i = m.group(0), m.end()
        if t == M.粗开:
            粗 = True
        elif t == M.粗关:
            粗 = False
        elif t == M.斜开:
            斜 = True
        elif t == M.斜关:
            斜 = False
        elif t == M.上开:
            上 = True
        elif t == M.上关:
            上 = False
        elif t == M.红开:
            红 = True
        elif t == M.红关:
            红 = False
        elif t.startswith("⟦m"):
            if 用数学 >= len(旧数学):
                raise ValueError("新文字里的行内公式比原来多，不能重建（公式只认 md）")
            p.append(旧数学[用数学])
            用数学 += 1
        else:
            if 用字段 >= len(旧字段):
                raise ValueError("新文字里的引用比原来多，不能重建（引用只认 md）")
            for r in 旧字段[用字段]:
                p.append(r)
            用字段 += 1
    if 缓:
        新run(缓, 粗, 斜, 上, 红)
    for b in 关:
        p.append(b)
    if 用数学 != len(旧数学) or 用字段 != len(旧字段):
        raise ValueError("公式/引用的个数对不上：公式 %d/%d，引用 %d/%d"
                         % (用数学, len(旧数学), 用字段, len(旧字段)))
    if ppr is not None and list(p).index(ppr) != 0:
        p.remove(ppr)
        p.insert(0, ppr)
