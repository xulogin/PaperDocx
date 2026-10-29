# -*- coding: utf-8 -*-
r"""出Word.py 之后再跑一遍：合并表格单元格、两行表头、列宽重分、⟦…⟧ 标红。

    py -3.12 -X utf8 排版/后处理.py 输入.docx 输出.docx      # 输入输出可以是同一个文件

出Word.py 的三线表只认一行表头、不合并单元格。md 里按下面的写法写，这里补上：

1. **两行表头**：md 表格第一行写组名，组名后面的格留空；第二行写子列名，
   第一行有名、第二行空的列（如「Dataset」）上下合并。组名横向合并、居中，
   组名下面加一条细线（cmidrule），子列名下面才是表头下线。
       | Variable | Dahurian larch |  |  |  | Asian white birch |  |  |  |
       |---|--:|--:|--:|--:|--:|--:|--:|--:|
       |  | Generated | Inventory | Bias | CCC | Generated | Inventory | Bias | CCC |
2. **分组表**：首列只在每组第一行写名字，下面留空 → 和上面那格竖向合并、垂直居中。
3. **列宽重分**：数值格放不下（被挤成两行）、或表头的词被拆开（「MAE」→「MA/E」）时，
   按格里实际要的宽度重分，总宽不变；文字列允许折行，但不拆单词（连字符处可断）。
   两行表头的表不重分（实测会把 Generated/real 拆开），自动分不开的用 `排版配置.表格["手定列宽"]`。
4. **标红**：要作者确认的地方在 md 里写成 ⟦…⟧，这里去掉括号、把里面的字标红（C00000）。

★ 2026-10-30 为一篇 RSE 稿子（6 张正文表 + 6 张补充表）实做出来的，每种情况都用真 Word
  渲染逐页看过。三条教训：
  - 表头箭头「MMD-CD ↓」：md 里箭头前写不换行空格（U+00A0），否则箭头单独掉到下一行。
  - 合并后的格要按「上边沿用首格、下边沿用末格」重设边线，否则表底粗线会丢。
  - python-docx 的 merge() 会把被并格的空段落搬过来，不删就是一格里几个空行。
"""
import copy
from functools import lru_cache
import re
import sys
from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import RGBColor

上线, 中线, 细线 = 12, 6, 4


def 边(tc, top=None, bottom=None):
    pr = tc.get_or_add_tcPr()
    for old in pr.findall(qn("w:tcBorders")):
        pr.remove(old)
    el = OxmlElement("w:tcBorders")
    for 名, 粗 in (("top", top), ("left", None), ("bottom", bottom), ("right", None)):
        b = OxmlElement("w:" + 名)
        if 粗:
            b.set(qn("w:val"), "single"); b.set(qn("w:sz"), str(粗))
            b.set(qn("w:space"), "0"); b.set(qn("w:color"), "000000")
        else:
            b.set(qn("w:val"), "nil")
        el.append(b)
    # tcBorders 要排在 tcW / gridSpan / vMerge 之后、shd / vAlign 之前
    后面 = [pr.find(qn(x)) for x in ("w:shd", "w:noWrap", "w:tcMar", "w:textDirection", "w:vAlign")]
    后面 = [x for x in 后面 if x is not None]
    if 后面:
        后面[0].addprevious(el)
    else:
        pr.append(el)


def 清空段(cell):
    """merge() 会把被并格的空段落搬过来，删掉多余的空段。"""
    ps = cell.paragraphs
    for p in ps[1:]:
        if not p.text.strip() and not p._p.findall(".//" + qn("m:oMath")):
            p._p.getparent().remove(p._p)


def 设表头样式(row, 表头行):
    trPr = row._tr.get_or_add_trPr()
    if 表头行 and trPr.find(qn("w:tblHeader")) is None:
        trPr.append(OxmlElement("w:tblHeader"))
    for c in row.cells:
        for p in c.paragraphs:
            p.style = "Table Header"


def 文(tbl, r, c):
    return tbl.rows[r].cells[c].text.strip()


def 处理两行表头(tbl):
    n = len(tbl.columns)
    if len(tbl.rows) < 3:
        return False
    r0 = [tbl.rows[0]._tr.tc_lst[i] for i in range(n)]
    t0 = [文(tbl, 0, i) for i in range(n)]
    t1 = [文(tbl, 1, i) for i in range(n)]
    # 判据：第一行有空格（组名后面留空），且第二行首格是空的
    if "" not in t0[1:] or t1[0] != "":
        return False
    末 = len(tbl.rows) - 1
    # 先定组：第一行非空格 + 其后连续的空格
    组 = []
    i = 0
    while i < n:
        j = i + 1
        while j < n and t0[j] == "":
            j += 1
        组.append((i, j))          # [i, j)
        i = j
    # 边线（合并前按格设，合并后保留首格的设置）
    for ci in range(n):
        单列 = any(a == ci and b - a == 1 for a, b in 组) and t1[ci] == ""
        边(tbl.rows[0].cells[ci]._tc, top=上线, bottom=中线 if 单列 else 细线)
        边(tbl.rows[1].cells[ci]._tc, top=None, bottom=中线)
    设表头样式(tbl.rows[0], True)
    设表头样式(tbl.rows[1], True)
    # 竖向合并：第一行有名、第二行空
    for a, b in 组:
        if b - a == 1 and t1[a] == "":
            m = tbl.cell(0, a).merge(tbl.cell(1, a))
            清空段(m)
            边(m._tc, top=上线, bottom=中线)
    # 横向合并：组名跨列（倒着来，免得列号变）
    for a, b in reversed(组):
        if b - a > 1:
            m = tbl.cell(0, a).merge(tbl.cell(0, b - 1))
            清空段(m)
            边(m._tc, top=上线, bottom=细线)
    return True


def 处理分组列(tbl, 起行):
    """首列留空的格和上面那格竖向合并。"""
    n行 = len(tbl.rows)
    col = [文(tbl, r, 0) for r in range(n行)]
    if "" not in col[起行 + 1:]:
        return 0
    合并数 = 0
    r = 起行
    while r < n行:
        if col[r] == "":
            r += 1
            continue
        s = r
        e = r
        while e + 1 < n行 and col[e + 1] == "":
            e += 1
        if e > s:
            上 = tbl.cell(s, 0)._tc
            保边 = copy.deepcopy(上.tcPr.find(qn("w:tcBorders"))) if 上.tcPr is not None else None
            下边 = tbl.cell(e, 0)._tc.tcPr.find(qn("w:tcBorders")) if tbl.cell(e, 0)._tc.tcPr is not None else None
            下底 = None if 下边 is None else 下边.find(qn("w:bottom"))
            m = tbl.cell(s, 0).merge(tbl.cell(e, 0))
            清空段(m)
            # 合并后的格：上边沿用首格，下边沿用末格（末行就是表底粗线）
            top = bottom = None
            if 保边 is not None and 保边.find(qn("w:top")).get(qn("w:val")) == "single":
                top = int(保边.find(qn("w:top")).get(qn("w:sz")))
            if 下底 is not None and 下底.get(qn("w:val")) == "single":
                bottom = int(下底.get(qn("w:sz")))
            边(m._tc, top=top, bottom=bottom)
            合并数 += 1
        r = e + 1
    return 合并数


_字体 = {}


@lru_cache(maxsize=None)
def _字宽cm(文, 号=9.0, 粗=False):
    """Times New Roman 字串的真实宽度（cm）：用 PIL 读本机字体的字形宽度，不靠估。

    ★ 2026-10-30 实测：按 0.5 em 估，粗体表头「MAE」「RMSE」估小了，排出来被拆成「MA/E」。
      找不到字体文件时退回粗估。
    """
    键 = 粗
    if 键 not in _字体:
        try:
            from PIL import ImageFont
            import os
            名 = "timesbd.ttf" if 粗 else "times.ttf"
            _字体[键] = ImageFont.truetype(os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", 名), 200)
        except Exception:
            _字体[键] = None
    f = _字体[键]
    if f is not None:
        return f.getlength(文) / 200 * 号 / 72 * 2.54
    em = 号 / 72 * 2.54
    w = 0
    for ch in 文:
        if ch.isupper() or ch in "—–±×%":
            w += 0.68
        elif ch in " .,;:()[]'":
            w += 0.28
        else:
            w += 0.5
    return w * em * (1.06 if 粗 else 1.0)


def _是数值格(文):
    """「0.185 ± 0.070」「0.32 [0.18, 0.44]」「+13.0%」「—」这种：不许折行。"""
    t = re.sub(r"\$[^$]*\$", "", 文)
    return bool(t) and not re.search(r"[A-Za-z]{2,}", t) and bool(re.search(r"[\d—]", t))


def _格边距cm(tbl):
    pr = tbl._tbl.tblPr
    mar = pr.find(qn("w:tblCellMar")) if pr is not None else None
    tw = 0
    if mar is not None:
        for 名 in ("left", "right"):
            e = mar.find(qn("w:" + 名))
            if e is not None:
                tw += int(e.get(qn("w:w")))
    return (tw or 170) / 567


def _表字号(tbl):
    for tc in tbl._tbl.iter(qn("w:tc")):
        for sz in tc.iter(qn("w:sz")):
            return int(sz.get(qn("w:val"))) / 2
    return 9.0


@lru_cache(maxsize=None)
def _行数_缓存(文, 宽, 号, 粗):
    return _行数_算(文, 宽, 号, 粗)


def _行数(文, 宽, 号, 粗):
    return _行数_缓存(文, round(宽, 3), 号, 粗)


def _行数_算(文, 宽, 号, 粗):
    """按词折行估一格要几行。词放不下时在连字符/斜杠后断；还放不下就硬拆（记一个大罚分）。"""
    if not 文:
        return 1
    if _是数值格(文):
        段 = [x.strip() for x in re.split(r"\s(?=\[)", 文)]
        if any(_字宽cm(x, 号, 粗) > 宽 for x in 段):
            return 99
        return 1 if _字宽cm(文, 号, 粗) <= 宽 else len(段)
    空格 = _字宽cm(" ", 号, 粗)
    行, 当前 = 1, 0.0
    for 词 in 文.split():
        片 = [p for p in re.split(r"(?<=[a-z]-)(?=[a-z])", 词) if p]
        for j, p in enumerate(片):
            w = _字宽cm(p, 号, 粗)
            if w > 宽:
                return 99                          # 这一格会被拆字
            加 = w + (空格 if (当前 and j == 0) else 0)
            if 当前 + 加 <= 宽:
                当前 += 加
            else:
                行 += 1
                当前 = w
    return 行


def _格们(tbl, 表头行数):
    """-> [(行, 起列, 跨列, 文, 粗, tc)]，竖向合并的续格（没字）不算。"""
    格 = []
    for ri, row in enumerate(tbl.rows):
        ci = 0
        for tc in row._tr.tc_lst:
            pr = tc.tcPr
            span = 1
            if pr is not None and pr.find(qn("w:gridSpan")) is not None:
                span = int(pr.find(qn("w:gridSpan")).get(qn("w:val")))
            文 = " ".join("".join(t.text or "" for t in p.iter(qn("w:t")))
                         for p in tc.iter(qn("w:p"))).strip()
            if 文:
                格.append((ri, ci, span, 文, ri < 表头行数, tc))
            ci += span
    return 格


def _数值格不断(tc):
    """数值格里「±」两边和方括号里的空格换成不换行空格：Word 只能在「[」前换行，和列宽模型一致。"""
    for t in tc.iter(qn("w:t")):
        x = t.text or ""
        y = re.sub(r" ?± ?", " ± ", x)
        y = re.sub(r"\[([^\]]*)\]", lambda m: "[" + m.group(1).replace(" ", " ") + "]", y)
        if y != x:
            t.text = y


def _排一种字号(格, n, 总, 边, 号):
    """给定字号，算最优列宽。-> (宽, 各行行数) 或 None（怎么分都放不下）"""
    def 下限():
        下 = [0.6] * n
        for ri, ci, span, 文, 粗, _ in 格:
            if span != 1:
                continue
            if _是数值格(文):
                需 = max(_字宽cm(p.strip(), 号, 粗) for p in re.split(r"\s(?=\[)", 文))
            else:
                片 = [p for 词 in 文.split() for p in re.split(r"(?<=[a-z]-)(?=[a-z])", 词) if p]
                需 = max(_字宽cm(p, 号, 粗) for p in 片)
            下[ci] = max(下[ci], 需 + 边)
        return 下

    def 行们(宽):
        行高 = {}
        for ri, ci, span, 文, 粗, _ in 格:
            w = sum(宽[ci:ci + span]) - 边
            行高[ri] = max(行高.get(ri, 1), _行数(文, w, 号, 粗))
        return 行高

    头行 = {ri for ri, ci, span, 文, 粗, _ in 格 if 粗}

    def 代价(宽):                                   # 表头行只算半个权重：先保正文不折行
        return sum(v * (0.5 if ri in 头行 else 1.0) for ri, v in 行们(宽).items())

    宽 = 下限()
    if sum(宽) > 总:
        return None
    步 = 0.05
    当前 = 代价(宽)
    while sum(宽) + 步 <= 总 + 1e-9:
        剩 = 总 - sum(宽)
        最好 = None
        for i in range(n):
            加 = 步
            while 加 <= 剩 + 1e-9:
                试 = 宽[:]
                试[i] += 加
                降 = 当前 - 代价(试)
                if 降 > 0:
                    率 = 降 / 加
                    if 最好 is None or 率 > 最好[0] + 1e-9:
                        最好 = (率, i, 加)
                    break
                加 += 步
        if 最好 is None:
            break
        _, i, 加 = 最好
        宽[i] += 加
        当前 = 代价(宽)
    # 剩下的宽：先补给还在折行的文字列（按缺口），都不折了再按比例摊
    剩 = 总 - sum(宽)
    if 剩 > 0:
        自然 = [0.0] * n
        for ri, ci, span, 文, 粗, _ in 格:
            if span == 1 and not 粗:                   # 只看正文格：剩下的宽先给正文
                自然[ci] = max(自然[ci], _字宽cm(文, 号, 粗) + 边)
        缺 = [max(0.0, 自然[i] - 宽[i]) for i in range(n)]
        给 = [剩 * c / sum(缺) for c in 缺] if sum(缺) > 0 else [剩 * w / sum(宽) for w in 宽]
        宽 = [宽[i] + 给[i] for i in range(n)]
    return 宽, 行们(宽)


def 优化列宽(tbl, 表头行数):
    """列宽 + 字号一起定：总宽不变，让整张表折行最少、最矮。

    ① 每列先给到「放得下最长不可断片段」的宽度（数值格整串不拆，只许在「[区间]」前换行；
       文字在空格处可断；连字符只在两边都是小写词时可断（height-to-diameter），
       FOR-species20K、BioDiv-3DTrees、1-NNA-CD 这类专名整体放下）；
    ② 剩下的宽一点一点分：每次给「加宽后总行数降得最多（按每厘米算）」的那一列；
    ③ 字号：从 9 pt 起，取「不拆字放得下」的最大字号（9 → 8.5 → 8），看得清优先，全文尽量一致。
       出Word.py 为了塞下宽表先降过字号的（如 11 列的表 4 被降到 8 pt），这里能放下就提回 9 pt。
    ★ 字宽用 PIL 量本机 Times New Roman 的真实字形（粗体表头单独量），不靠估。
    """
    grid = tbl._tbl.find(qn("w:tblGrid"))
    if grid is None:
        return False
    cols = grid.findall(qn("w:gridCol"))
    n = len(cols)
    旧 = [int(c.get(qn("w:w"))) / 567 for c in cols]
    总 = sum(旧)
    边 = _格边距cm(tbl) + 0.06                       # 字宽是真实量的，只留一点舍入余量
    原号 = _表字号(tbl)
    格 = _格们(tbl, 表头行数)
    if not 格:
        return False

    def 表高(号, 行高):                               # pt：行数 × 行距 + 每行段前段后
        return sum(行高.values()) * 号 * 1.15 + len(行高) * 3.0

    方案 = {}
    for 号 in (9.0, 8.5, 8.0):                      # 看得清优先：从 9 pt 起，取放得下的最大字号
        r = _排一种字号(格, n, 总, 边, 号)
        if r:
            方案[号] = (r[0], 表高(号, r[1]), max(r[1].values()))
            break
    if not 方案:
        return False
    号 = next(iter(方案))
    宽 = 方案[号][0]
    缩号 = 号 if 号 != 原号 else None

    for ri, ci, span, 文, 粗, tc in 格:
        if _是数值格(文):
            _数值格不断(tc)
    if 缩号 is not None:
        for r in tbl._tbl.iter(qn("w:r")):
            rpr = r.find(qn("w:rPr"))
            if rpr is None:
                rpr = OxmlElement("w:rPr"); r.insert(0, rpr)
            for 名 in ("w:sz", "w:szCs"):
                for old in rpr.findall(qn(名)):
                    rpr.remove(old)
                e = OxmlElement(名); e.set(qn("w:val"), str(int(缩号 * 2))); rpr.append(e)
        for r in tbl._tbl.iter(qn("m:r")):
            rpr = r.find(qn("w:rPr"))
            if rpr is not None:
                for 名 in ("w:sz", "w:szCs"):
                    for old in rpr.findall(qn(名)):
                        old.set(qn("w:val"), str(int(缩号 * 2)))
    for c, w in zip(cols, 宽):
        c.set(qn("w:w"), str(int(w * 567)))
    for row in tbl.rows:
        ci = 0
        for tc in row._tr.tc_lst:
            pr = tc.get_or_add_tcPr()
            span = 1
            if pr.find(qn("w:gridSpan")) is not None:
                span = int(pr.find(qn("w:gridSpan")).get(qn("w:val")))
            tcw = pr.find(qn("w:tcW"))
            if tcw is not None:
                tcw.set(qn("w:w"), str(int(sum(宽[ci:ci + span]) * 567)))
                tcw.set(qn("w:type"), "dxa")
            ci += span
    return f"{号:g}pt"


def _手定列宽():
    """自动分不开的表（十几列、表头词又长），在项目的 排版配置.表格["手定列宽"] 里按表头签名写死列宽。

    键是表头前几格的文字（tuple），值是每列宽度（cm，按比例拉满正文宽）：
        手定列宽={("Species", "Reference", "Variable"): [1.4, 2.05, 1.75, ...]}
    """
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    try:
        import 排版配置 as C
        return dict(C.表格.get("手定列宽", {}) or {})
    except Exception:
        return {}


def 套手定列宽(tbl):
    头 = tuple(c.text.strip() for c in tbl.rows[0].cells)
    for 签, 列 in _手定列宽().items():
        if 头[:len(签)] == 签 and len(头) == len(列):
            grid = tbl._tbl.find(qn("w:tblGrid"))
            cols = grid.findall(qn("w:gridCol"))
            总 = sum(int(c.get(qn("w:w"))) for c in cols) / 567
            新 = [w * 总 / sum(列) for w in 列]
            for c, w in zip(cols, 新):
                c.set(qn("w:w"), str(int(w * 567)))
            for row in tbl.rows:
                ci = 0
                for tc in row._tr.tc_lst:
                    pr = tc.get_or_add_tcPr()
                    span = 1
                    if pr.find(qn("w:gridSpan")) is not None:
                        span = int(pr.find(qn("w:gridSpan")).get(qn("w:val")))
                    tcw = pr.find(qn("w:tcW"))
                    if tcw is not None:
                        tcw.set(qn("w:w"), str(int(sum(新[ci:ci + span]) * 567)))
                    ci += span
            return True
    return False


def 标红(doc):
    红 = RGBColor(0xC0, 0x00, 0x00)
    计 = 0
    段们 = list(doc.paragraphs)
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                段们 += c.paragraphs
    for p in 段们:
        if "⟦" not in p.text and "⟧" not in p.text:
            continue
        在内 = False
        for r in list(p._p.iter(qn("w:r"))):
            ts = r.findall(qn("w:t"))
            if not ts:
                if 在内:
                    _设红(r)
                continue
            全 = "".join(t.text or "" for t in ts)
            if "⟦" not in 全 and "⟧" not in 全:
                if 在内:
                    _设红(r)
                continue
            # 把这个 run 按标记切开
            段 = []
            buf = ""
            for ch in 全:
                if ch == "⟦":
                    段.append((buf, 在内)); buf = ""; 在内 = True; 计 += 1
                elif ch == "⟧":
                    段.append((buf, 在内)); buf = ""; 在内 = False
                else:
                    buf += ch
            段.append((buf, 在内))
            for 字, 红否 in 段:
                if not 字:
                    continue
                新 = copy.deepcopy(r)
                for t in 新.findall(qn("w:t")):
                    新.remove(t)
                t = OxmlElement("w:t")
                t.text = 字
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                新.append(t)
                if 红否:
                    _设红(新)
                r.addprevious(新)
            r.getparent().remove(r)
    return 计


def _设红(r):
    rpr = r.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        r.insert(0, rpr)
    for old in rpr.findall(qn("w:color")):
        rpr.remove(old)
    c = OxmlElement("w:color")
    c.set(qn("w:val"), "C00000")
    rpr.append(c)


def main(src, dst):
    doc = Document(src)
    for k, t in enumerate(doc.tables):
        if len(t.columns) < 2 or len(t.rows) < 2:
            continue                    # 公式双列表这类不碰
        if not any(p.style.name in ("Table Header", "Table Cell")
                   for row in t.rows for c in row.cells for p in c.paragraphs):
            continue
        两行 = 处理两行表头(t)
        并 = 处理分组列(t, 2 if 两行 else 1)
        宽 = 套手定列宽(t) or 优化列宽(t, 2 if 两行 else 1)
        if 两行 or 并 or 宽:
            print(f"  表 #{k + 1}（首格「{t.rows[0].cells[0].text.strip()[:20]}」）：两行表头={两行}，首列合并 {并} 组，重分列宽={宽}")
    n = 标红(doc)
    print(f"  标红 {n} 处")
    doc.save(dst)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2])
