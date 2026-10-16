# -*- coding: utf-8 -*-
r"""三件套共用的**块模型** —— 出Word.py 和 三方同步.py 都从这里拿解析结果，
只有一份实现，不会两边漂移。

## 三件套

    C_英文稿V41_英文版.md        纯英文（派生，但也是同步的一方）
    C_英文稿V41_中英文对照版.md   英文 + 中文（作者的读取界面）
    C_英文稿V41_排版版.docx       Word 排版版（样式归作者，文字联动）

英文**块序列**是三者的公共骨架：对照版去掉 `**中文译文：**` 和译文块就是英文版，
docx 里每个块起始处埋一个书签 `PDX0001`（前缀见 排版配置.锚点前缀）。三方同步就是比这三份的块文本。

## 块类型（扫 V40 实测，共 12 种，没有别的）

    h1 h2 h3 h4   标题
    para          正文段
    keywords      `**Keywords:** ...`
    image         `![](figures/xxx.png){width=16cm}`
    figcap        `**Figure 1.** ...`
    tabcap        `**Table 1. ...**`（整句被 ** 包住，排版时只把 "Table 1." 加粗）
    table         markdown 管道表
    tabnote       `Note: ...`
    equation      `$$ ... \qquad\qquad (n) $$` + 可选 `<!--eq:xxx-->` 锚

## 为什么要「可同步文本」这个概念

Word 里的一个段落是若干 run，md 里是一串带标记的文本。要双向对，得有个
**公共表示**：把行内公式换成占位符 `⟦m1⟧`，保留 `**粗**` / `*斜*`，
引用保留成渲染文本 `(Wang et al., 2022)`（docx 里它是 Zotero 字段，
取字段的显示文本就正好等于这个）。

★ 公式和引用**只认 md 为准**：在 Word 里改公式 / 改引用文本，同步会报错不写，
  让人回 md 改。理由是 OMML 反推 LaTeX 不可靠，宁可报错也不许猜。
"""
from __future__ import annotations

import re
from pathlib import Path

CJK = re.compile(r"[一-鿿㐀-䶿豈-﫿぀-ヿ]")
译文标记 = "**中文译文：**"


# ============================================================ 切块
def 切块(正文: str):
    """按空行切块，返回 list[list[str]]（每块是若干行）。"""
    正文 = 正文.replace("\r\n", "\n")
    出, 当前 = [], []
    for 行 in 正文.split("\n"):
        if 行.strip() == "":
            if 当前:
                出.append(当前)
                当前 = []
        else:
            当前.append(行)
    if 当前:
        出.append(当前)
    return 出


def 认类型(块行) -> str:
    首 = 块行[0].strip()
    m = re.match(r"^(#{1,4}) ", 首)
    if m:
        return "h%d" % len(m.group(1))
    if 首.startswith("![]("):
        return "image"
    if 首.startswith("$$"):
        return "equation"
    if 首.startswith("|"):
        return "table"
    if 首.startswith("**Keywords:**"):
        return "keywords"
    if 首.startswith("**Authors:**"):
        return "authors"
    if 首.startswith("**Affiliations:**"):
        return "affil"
    if 首.startswith("**Correspondence:**"):
        return "corresp"
    if re.match(r"^\*\*Figure [0-9A]+\.", 首):
        return "figcap"
    if re.match(r"^\*\*Table [0-9A]+\.", 首):
        return "tabcap"
    if 首.startswith("Note:"):
        return "tabnote"
    if 首.startswith("(Reference list is inserted"):
        # 参考文献占位块：出 Word 时整段被 Zotero 书目字段顶替掉，
        # 所以它在 docx 里没有对应段落，不参与三方同步。
        return "bibph"
    return "para"


class 块:
    __slots__ = ("序", "类型", "文本", "中文")

    def __init__(self, 序, 类型, 文本, 中文=None):
        self.序 = 序              # 1 起的块号，就是书签 <前缀>%04d 的号
        self.类型 = 类型
        self.文本 = 文本          # md 原文（多行用 \n 连）
        self.中文 = 中文          # 只有对照版才有

    @property
    def 锚(self):
        # ★ 前缀走配置（排版配置.锚点前缀），不写死 —— 换项目就能换一套书签名
        import 排版配置
        return "%s%04d" % (排版配置.锚点前缀, self.序)

    def __repr__(self):
        return f"<块{self.序} {self.类型} {self.文本[:40]!r}>"


def 读英文版(路径) -> list:
    """纯英文 md → [块]"""
    s = Path(路径).read_text(encoding="utf-8")
    out = []
    for i, b in enumerate(切块(s), 1):
        t = "\n".join(b)
        out.append(块(i, 认类型(b), t))
    return out


def 读对照版(路径) -> list:
    r"""中英文对照版 md → [块]（每块带 .中文）

    结构（`_工作区/出英文版.py` 已验证过的）：

        英文块 / 空行 / **中文译文：** / 空行 / 中文块

    标题是例外：英文标题块紧跟中文标题块，**没有** `**中文译文：**` 标记。
    图片块没有译文。
    """
    s = Path(路径).read_text(encoding="utf-8")
    块们 = 切块(s)
    是标题 = lambda b: bool(re.match(r"#{1,4} ", b[0]))
    out, i, 序 = [], 0, 0
    while i < len(块们):
        b = 块们[i]
        if b[0].strip() == 译文标记:
            raise SystemExit(f"⛔ 对照版第 {i} 块是孤立的 `{译文标记}`，结构坏了")
        序 += 1
        英 = "\n".join(b)
        中 = None
        if 是标题(b) and i + 1 < len(块们) and 是标题(块们[i + 1]):
            中 = "\n".join(块们[i + 1])
            i += 2
        elif (i + 2 < len(块们) and 块们[i + 1][0].strip() == 译文标记):
            中 = "\n".join(块们[i + 2])
            i += 3
        else:
            i += 1
        out.append(块(序, 认类型(b), 英, 中))
    return out


def 写对照版(路径, 块们):
    """[块] → 对照版 md（CRLF，和既有文件一致）"""
    段 = []
    for b in 块们:
        段.append(b.文本)
        if b.中文 is None:
            continue
        if b.类型.startswith("h"):
            段.append(b.中文)                 # 标题：紧跟，无标记
        else:
            段.append(译文标记)
            段.append(b.中文)
    文 = "\n\n".join(段) + "\n"
    Path(路径).write_bytes(文.replace("\n", "\r\n").encode("utf-8"))


def 写英文版(路径, 块们):
    文 = "\n\n".join(b.文本 for b in 块们) + "\n"
    Path(路径).write_bytes(文.replace("\n", "\r\n").encode("utf-8"))


# ============================================================ 可同步文本
# 把一块 md 规约成「可和 Word 段落逐字比对」的形式。
# 规则：行内 $...$ → ⟦m1⟧⟦m2⟧…（按出现顺序编号），其余原样。
行内公式 = re.compile(r"(?<!\$)\$(?!\$)((?:[^$\n]|\\\$)+?)\$(?!\$)")


def 抽公式(文本: str):
    """'a $x$ b $y$' → ('a ⟦m1⟧ b ⟦m2⟧', ['x', 'y'])"""
    拿 = []

    def 换(m):
        拿.append(m.group(1))
        return "⟦m%d⟧" % len(拿)

    return 行内公式.sub(换, 文本), 拿


def 填公式(骨架: str, 公式们):
    def 换(m):
        i = int(m.group(1)) - 1
        if i >= len(公式们):
            raise ValueError("占位符 ⟦m%s⟧ 超出公式表长度" % m.group(1))
        return "$" + 公式们[i] + "$"

    return re.sub(r"⟦m(\d+)⟧", 换, 骨架)


def 规约(文本: str) -> str:
    """只用来比「变没变」：去掉行内公式内容、压掉多余空白。"""
    骨, _ = 抽公式(文本)
    return re.sub(r"[ \t]+", " ", 骨).strip()


# ============================================================ 同步串
# md 的一段和 Word 的一段要能逐字比，得先化成同一种**中间串**：
#
#     **粗** *斜* ^上标^ $x$ (Wang et al., 2022)
#   → ⟪b⟫粗⟪/b⟫ ⟪i⟫斜⟪/i⟫ ⟪s⟫上标⟪/s⟫ ⟦m1⟧ ⟦c:ref1⟧
#
# ★ 为什么哨兵用 ⟪⟫ 而不是直接留 `**`：稿子里有真的星号（作者标记 `^1,\*^`），
#   留 `**` 会和真星号撞车，反过来写回 md 时分不清哪个该转义。
粗开, 粗关 = "⟪b⟫", "⟪/b⟫"
斜开, 斜关 = "⟪i⟫", "⟪/i⟫"
上开, 上关 = "⟪s⟫", "⟪/s⟫"
# 红：只在「和上一版 diff 标红」时用，md 里**不存在**这对哨兵，
# 它只活在内存里，从 md 生成同步串 → 打红标记 → 写进 docx 这一条路上。
红开, 红关 = "⟪r⟫", "⟪/r⟫"
哨兵 = re.compile(r"⟪/?[bisr]⟫|⟦m\d+⟧|⟦c:[^⟧]*⟧")


def md转同步串(文本: str, 引用反查=None):
    r"""md 行内 → 同步串。返回 (串, 公式 LaTeX 列表, 引用 refN 组列表)。

    `引用反查` 是 {渲染文本: refN}；给了才把 `(Wang et al., 2022)` 收成 ⟦c:ref1⟧。
    """
    公式, 引用 = [], []

    def 吃公式(m):
        公式.append(m.group(1))
        return "⟦m%d⟧" % len(公式)

    s = 行内公式.sub(吃公式, 文本)

    if 引用反查 is not None:
        def 吃引用(m):
            部分 = [p.strip() for p in m.group(0)[1:-1].split(";")]
            if not all(p in 引用反查 for p in 部分):
                return m.group(0)
            引用.append([引用反查[p] for p in 部分])
            return "⟦c:%s⟧" % ";".join(引用[-1])

        s = 引用反查["__正则__"].sub(吃引用, s)

    out, i, n = [], 0, len(s)
    while i < n:
        c = s[i]
        if c == "\\" and i + 1 < n:          # \* \_ \$ 之类的转义，脱一层
            out.append(s[i + 1])
            i += 2
        elif s.startswith("**", i):
            j = s.find("**", i + 2)
            if j < 0:
                out.append("**")
                i += 2
            else:
                out.append(粗开 + s[i + 2:j].replace("\\", "") + 粗关)
                i = j + 2
        elif c == "*":
            j = s.find("*", i + 1)
            if j < 0:
                out.append("*")
                i += 1
            else:
                out.append(斜开 + s[i + 1:j].replace("\\", "") + 斜关)
                i = j + 1
        elif c == "^":
            j = s.find("^", i + 1)
            if j < 0 or j - i > 24:          # 孤零零的 ^（比如 m⁻² 旁边）不当上标
                out.append("^")
                i += 1
            else:
                out.append(上开 + s[i + 1:j].replace("\\", "") + 上关)
                i = j + 1
        else:
            out.append(c)
            i += 1
    return "".join(out), 公式, 引用


def 同步串转md(串: str, 公式, 引用渲染=None) -> str:
    """同步串 → md 行内。公式按序填回；引用按 refN 还原成渲染文本。"""
    def 转义(t):
        # 真星号要转义，否则下一轮解析会当成强调
        return t.replace("*", "\\*")

    out, i, n = [], 0, len(串)
    while i < n:
        m = 哨兵.match(串, i)
        if not m:
            j = 哨兵.search(串, i)
            末 = j.start() if j else n
            out.append(转义(串[i:末]))
            i = 末
            continue
        t = m.group(0)
        i = m.end()
        if t == 粗开:
            j = 串.index(粗关, i)
            out.append("**" + 转义(串[i:j]) + "**")
            i = j + len(粗关)
        elif t == 斜开:
            j = 串.index(斜关, i)
            out.append("*" + 转义(串[i:j]) + "*")
            i = j + len(斜关)
        elif t == 上开:
            j = 串.index(上关, i)
            out.append("^" + 转义(串[i:j]) + "^")
            i = j + len(上关)
        elif t.startswith("⟦m"):
            k = int(t[2:-1]) - 1
            if k >= len(公式):
                raise ValueError("公式占位 %s 没有对应的 LaTeX" % t)
            out.append("$" + 公式[k] + "$")
        else:                                  # ⟦c:ref1;ref2⟧
            refs = t[3:-1].split(";")
            if 引用渲染 is None:
                out.append(t)
            else:
                out.append("(" + "; ".join(引用渲染[r] for r in refs) + ")")
    return "".join(out)


def 压空白(s: str) -> str:
    return re.sub(r"[ \t]+", " ", s).strip()


# ============================================================ 表格
def 解析表(块文本: str):
    """markdown 管道表 → (对齐列表, [[单元格]])，表头是第 0 行。"""
    行 = [l.strip() for l in 块文本.split("\n") if l.strip()]
    格 = []
    对齐 = None
    for idx, l in enumerate(行):
        cells = [c.strip() for c in l.strip("|").split("|")]
        if idx == 1 and all(re.fullmatch(r":?-{2,}:?", c) for c in cells):
            对齐 = ["center" if c.startswith(":") and c.endswith(":")
                   else "right" if c.endswith(":")
                   else "left" for c in cells]
            continue
        格.append(cells)
    if 对齐 is None:
        raise ValueError("表格缺分隔行：" + 块文本[:80])
    宽 = max(len(r) for r in 格)
    for r in 格:
        r += [""] * (宽 - len(r))
    对齐 += ["left"] * (宽 - len(对齐))
    return 对齐[:宽], 格


def 组表(对齐, 格) -> str:
    r"""反过来：(对齐, 格) → markdown 表（同步写回 md 时用）

    ★ 空单元格写成 `| |` 而不是 `|  |`（一个空格，不是两个）——
    稿子里本来就是这么写的，写成两个空格会让「表格往返一致」这道检查失败，
    于是每次同步都误报「表变了」。
    """
    线 = {"center": ":-:", "right": "--:", "left": "---"}
    一行 = lambda cs: "|" + "|".join(f" {c} " if c else " " for c in cs) + "|"
    行 = [一行(格[0]), "|" + "|".join(线[a] for a in 对齐) + "|"]
    for r in 格[1:]:
        行.append(一行(r))
    return "\n".join(行)


# ============================================================ 公式块
公式块 = re.compile(r"^\$\$\s*\n(.*?)\n\$\$\s*(?:\n<!--eq:([^>]*)-->)?\s*$", re.S)


def 拆公式块(块文本: str):
    """'$$\\n... \\qquad\\qquad (3)\\n$$\\n<!--eq:prior-->'
       → (latex 本体, '3', 'eq:prior' 或 None)"""
    m = 公式块.match(块文本.strip())
    if not m:
        raise ValueError("公式块认不出来：" + 块文本[:100])
    体, 锚 = m.group(1), m.group(2)
    n = re.search(r"\((S?\d+)\)\s*$", 体)
    号 = None
    if n:
        号 = n.group(1)
        体 = 体[: n.start()]
        体 = re.sub(r"(?:,?\s*\\qquad)+\s*$", "", 体)
    return 体.strip(), 号, 锚
