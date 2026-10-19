# -*- coding: utf-8 -*-
r"""三件套文字联动：英文版 md ↔ 中英文对照版 md ↔ 排版版 docx。

    py -3.12 -X utf8 _工作区/排版/三方同步.py            看一眼谁改了，不写文件
    py -3.12 -X utf8 _工作区/排版/三方同步.py --写        真的同步并写回
    py -3.12 -X utf8 _工作区/排版/三方同步.py --重置快照   以当前三份为准重新建基线

## 规矩

1. **只联动文字，样式一律不碰。** 作者在 Word 里调的字体、间距、缩进、
   分页、样式定义，同步一个都不动。
2. **三方任意一方改英文都能传到另外两方**；同一块被两方改成不同的内容
   → 报冲突、**一个字都不写**，让人决定。
3. **中文只住在对照版**。英文变了，这块的中文就记成「待跟」，
   写进 `_待跟中文.txt`，由 Claude 按 CLAUDE.md 铁律 1 的四步去改，不自动翻译。
4. **行内公式和引用只认 md。** 在 Word 里改了公式或引用，同步报错不写，
   回 md 改 —— OMML 反推 LaTeX 不可靠，宁可报错也不猜。
5. 图块、独立公式块在 Word 里不可编辑文字，不参与 docx 侧比对。

## 快照

`_同步快照.json` 记的是**上一次同步完成时**三方各自的样子。
谁和快照不一样，谁就是「改过的那一方」。所以：

  ⛔ **别手动删快照**，删了就分不清谁改的，只能 `--重置快照` 重新拉齐。
  ⛔ 重建 docx（`出Word.py`）之后要 `--重置快照`，否则会把重建当成「docx 改了」。
"""
from __future__ import annotations

import json
import re
import sys
from pathlib import Path

from docx import Document

这里 = Path(__file__).resolve().parent
sys.path.insert(0, str(这里))
import 块模型 as M                      # noqa: E402
import docx读写 as D                     # noqa: E402

# ★ 2026-10-16 的教训：版本号**不许写死**。V42 开出来之后这里还指着 V41，
#   同步报「275 块」（当前稿是 266），等于在同步一份已经冻结的旧稿。
import _项目                                          # noqa: E402
根, _版, 英文版, 对照版, 排版版 = _项目.三件套()
快照路径 = 这里 / "_同步快照.json"
待跟路径 = 这里 / "_待跟中文.txt"

# docx 侧能读写文字的块类型。image / equation 不在内：Word 里那是图和 OMML，
# 没有可编辑的纯文字，md 说了算。
可同步 = {"h1", "h2", "h3", "h4", "para", "keywords", "figcap", "tabcap",
        "tabnote", "authors", "affil", "corresp", "table"}
多段 = {"affil"}                         # 一个 md 块对应多个 Word 段落


# ══════════════════════════════════════════════════ 引用对照
def 载引用():
    渲染 = {}
    import 出Word
    映 = next(p for p in (根 / "引用映射.tsv", 根 / "_工作区/引用映射.tsv",
                         出Word.找引用数据() / "引用映射.tsv") if p.exists())
    for line in 映.read_text("utf-8").splitlines()[1:]:
        ref, 文 = line.split("\t")
        渲染[ref] = 文
    反查 = {v: k for k, v in 渲染.items()}
    形 = sorted(渲染.values(), key=len, reverse=True)
    单 = "|".join(re.escape(s) for s in 形)
    反查["__正则__"] = re.compile(r"\((%s)(?:; (?:%s))*\)" % (单, 单))
    import 出Word
    索引 = json.loads((出Word.找引用数据() / "zot_match_index.json").read_text("utf-8"))
    条目id表 = {v["itemKey"]: k for k, v in 索引.items()}
    return 渲染, 反查, 条目id表


# ══════════════════════════════════════════════════ 三方各自的「可比表示」
def md侧(块们, 反查):
    """{序: 可比值}。表格是 [[同步串]]，其余是同步串。"""
    出, 附 = {}, {}
    for b in 块们:
        if b.类型 not in 可同步:
            continue
        if b.类型 == "table":
            对齐, 格 = M.解析表(b.文本)
            串格 = [[M.压空白(M.md转同步串(c, 反查)[0]) for c in 行] for 行 in 格]
            出[b.序] = 串格
            附[b.序] = ("table", 对齐,
                       [[M.md转同步串(c, 反查) for c in 行] for 行 in 格])
        elif b.类型 == "affil":
            行 = [l.strip() for l in b.文本.split("\n")[1:] if l.strip()]
            出[b.序] = [M.压空白(M.md转同步串(l, 反查)[0]) for l in 行]
            附[b.序] = ("affil", None, [M.md转同步串(l, 反查) for l in 行])
        else:
            # md 带着 `## `、`**Authors:** ` 这些标记，Word 里没有（层级在样式上）。
            # 比之前先把标记摘掉存进「前缀」，写回 md 时再装回去。
            正 = b.文本
            前缀 = ""
            m = re.match(r"^(#{1,4} )", 正)
            if m:
                前缀, 正 = m.group(1), 正[m.end():]
            else:
                for p in ("**Authors:**", "**Correspondence:**"):
                    if 正.startswith(p):
                        前缀, 正 = p + " ", 正[len(p):].lstrip()
            串, 公式, 引用 = M.md转同步串(正, 反查)
            出[b.序] = M.压空白(串)
            附[b.序] = (b.类型, 前缀, (串, 公式, 引用))
    return 出, 附


def docx侧(doc, 块们, 条目id表):
    书签 = D.书签表(doc)
    出, 位置 = {}, {}
    缺 = []
    for b in 块们:
        if b.类型 not in 可同步:
            continue
        el = 书签.get(b.锚)
        if el is None:
            缺.append(b)
            continue
        if b.类型 == "table":
            出[b.序] = [[M.压空白(c) for c in 行] for 行 in D.表转格(el, 条目id表)]
            位置[b.序] = el
        elif b.类型 in 多段:
            n = len([l for l in b.文本.split("\n")[1:] if l.strip()])
            段 = D.段落跨度(doc, el, n)
            出[b.序] = [M.压空白(D.段转同步串(p, 条目id表)) for p in 段]
            位置[b.序] = 段
        else:
            出[b.序] = M.压空白(D.段转同步串(el, 条目id表))
            位置[b.序] = el
    return 出, 位置, 缺


# ══════════════════════════════════════════════════ 写回
def 写回md块(b, 新值, 附, 渲染):
    """把合并后的可比值写回一个 md 块的 .文本。"""
    类型, 附1, 原 = 附[b.序]
    if 类型 == "table":
        对齐 = 附1
        格 = [[M.同步串转md(c, 原[ri][ci][1], 渲染) for ci, c in enumerate(行)]
             for ri, 行 in enumerate(新值)]
        b.文本 = M.组表(对齐, 格)
    elif 类型 == "affil":
        行 = [M.同步串转md(s, 原[i][1], 渲染) for i, s in enumerate(新值)]
        b.文本 = "**Affiliations:**\n" + "\n".join(行)
    else:
        b.文本 = (附1 or "") + M.同步串转md(新值, 原[1], 渲染)


def 写回docx块(b, 新值, 位置, 条目id表):
    el = 位置[b.序]
    if b.类型 == "table":
        trs = el.findall(D.W + "tr")
        for ri, 行 in enumerate(新值):
            tcs = trs[ri].findall(D.W + "tc")
            for ci, 串 in enumerate(行):
                ps = tcs[ci].findall(D.W + "p")
                D.写段(ps[0], 串, 条目id表)
                for 多 in ps[1:]:
                    tcs[ci].remove(多)
    elif b.类型 in 多段:
        for p, 串 in zip(el, 新值):
            D.写段(p, 串, 条目id表)
    else:
        D.写段(el, 新值, 条目id表)


# ══════════════════════════════════════════════════ 主
def 主():
    写 = "--写" in sys.argv
    重置 = "--重置快照" in sys.argv
    渲染, 反查, 条目id表 = 载引用()

    # ★★ 派生模式挡在门口（2026-10-19 补）。
    #   作者手写的稿子里 `## References` 底下是几十条纯文本，`做排版源.py` 会派生出
    #   `<版>_排版源.md`（文献表换成 1 个 Zotero 占位块、叙述式引用改写过），
    #   **docx 是从排版源出的**。这时候拿 `英文版.md` 来对，块数直接差几十个，
    #   书签全错位 —— 轻则一堆假警报，重则把 A 块的文字写进 B 块。
    #   派生模式下的正确做法：文字改在 `英文版.md`，重跑 `走一遍.py`，不走同步。
    排版源 = 排版版.with_name(f"{_版}_排版源.md")
    if 排版源.exists():
        甲块 = M.读英文版(英文版)
        源块 = M.读英文版(排版源)
        if len(甲块) != len(源块):
            raise SystemExit(
                f"⛔ 这个项目是**派生模式**：docx 是从 {排版源.name} 出的，"
                f"不是从 {英文版.name}。\n"
                f"   {英文版.name} {len(甲块)} 块 ≠ {排版源.name} {len(源块)} 块"
                f"（差的是文献表那几十条纯文本），按书签对会整体错位。\n"
                "   ⇒ 派生模式**不要用三方同步**：文字改在英文版 md 里，"
                "重跑 `走一遍.py` 全量重建。\n"
                "   （作者在 Word 里改字这条路，只适用于文献表本来就是占位块的原生格式稿。）")

    甲 = M.读英文版(英文版)                      # A：英文版 md
    # ★ 对照版是**可选**的。只有「作者只读中文、要中英对照」的项目才有它；
    #   纯英文项目就只同步「英文版 ↔ docx」两方。
    有对照 = 对照版.exists()
    乙 = M.读对照版(对照版) if 有对照 else [
        M.块(b.序, b.类型, b.文本) for b in 甲]
    if len(甲) != len(乙):
        raise SystemExit(f"⛔ 英文版 {len(甲)} 块 ≠ 对照版 {len(乙)} 块。"
                         "有人在某一边加删了整段，同步认不回去，先人工对齐。")
    for a, b in zip(甲, 乙):
        if a.类型 != b.类型:
            raise SystemExit(f"⛔ 块 {a.序} 两边类型不同：{a.类型} / {b.类型}")

    doc = Document(排版版)
    甲值, 甲附 = md侧(甲, 反查)
    乙值, 乙附 = md侧(乙, 反查)
    丙值, 丙位, 缺 = docx侧(doc, 甲, 条目id表)
    if 缺:
        print(f"⚠ docx 里有 {len(缺)} 个块找不到书签锚点（整段被删过？）："
              f"{[b.锚 for b in 缺[:8]]}")
        print("  这些块这一轮按「docx 未改」处理，不会被当成删除。")

    快照 = {}
    if 快照路径.exists() and not 重置:
        快照 = json.loads(快照路径.read_text("utf-8")).get("块", {})
    首次 = not 快照

    变, 冲突, 中文待跟 = [], [], []
    for b in 甲:
        if b.类型 not in 可同步:
            continue
        k = b.锚
        旧 = 快照.get(k, {})
        候选 = []
        for 名, 值 in (("英文版", 甲值.get(b.序)), ("对照版", 乙值.get(b.序)),
                     ("Word", 丙值.get(b.序))):
            if 值 is None:
                continue
            if 首次 or 值 != 旧.get(名):
                候选.append((名, 值))
        不同 = {json.dumps(v, ensure_ascii=False) for _, v in 候选}
        if 首次:
            if len(不同) > 1:
                冲突.append((b, 候选, "首次建快照时三方就不一致"))
            continue
        if not 候选:
            continue
        if len(不同) > 1:
            冲突.append((b, 候选, "多方改成了不同内容"))
            continue
        变.append((b, 候选[0][0], 候选[0][1]))

    # 中文待跟：英文变了的块，它的中文就过期了
    for b, 源, 值 in 变:
        乙b = 乙[b.序 - 1]
        if 乙b.中文:
            中文待跟.append((b, 源))

    print(f"英文版/对照版/Word 共 {len(甲)} 块，可同步 {len(甲值)} 块")
    if 首次 and not 重置:
        print("（没有快照，这一轮只建基线）")
    if 冲突:
        print(f"\n⛔ 冲突 {len(冲突)} 处，**什么都没写**：")
        for b, 候选, 因 in 冲突[:12]:
            print(f"  块 {b.锚}（{b.类型}）{因}")
            for 名, 值 in 候选:
                s = 值 if isinstance(值, str) else json.dumps(值, ensure_ascii=False)
                print(f"      {名:6s} {s[:110]}")
        if len(冲突) > 12:
            print(f"  …… 还有 {len(冲突) - 12} 处")
        raise SystemExit(1)

    if not 变 and not 首次:
        print("✓ 三方一致，没有要同步的")
        return

    for b, 源, 值 in 变:
        s = 值 if isinstance(值, str) else json.dumps(值, ensure_ascii=False)
        print(f"  {b.锚} {b.类型:8s} ← {源:6s} {s[:90]}")
    print(f"\n共 {len(变)} 块要同步；中文待跟 {len(中文待跟)} 块")

    if not 写:
        print("\n（这是查看模式，没写文件。要真写加 --写）")
        return

    for b, 源, 值 in 变:
        if 源 != "英文版":
            写回md块(甲[b.序 - 1], 值, 甲附, 渲染)
        if 源 != "对照版":
            写回md块(乙[b.序 - 1], 值, 乙附, 渲染)
        if 源 != "Word" and b.序 in 丙位:
            写回docx块(b, 值, 丙位, 条目id表)

    M.写英文版(英文版, 甲)
    if 有对照:
        M.写对照版(对照版, 乙)
    doc.save(排版版)
    # python-docx 存盘后 docProps/custom.xml（Zotero 会话信息）不一定还在，
    # 掉了的话 Zotero 就不认这个文档了。无脑补一次，幂等。
    import 出Word
    出Word.装Zotero偏好(排版版)

    存快照(甲, 乙, 条目id表, 反查)
    if 中文待跟:
        待跟路径.write_text(
            "英文改过、中文还没跟上的块（按 CLAUDE.md 铁律 1 的四步处理）\n"
            + "\n".join(f"{b.锚}\t{b.类型}\t来源 {源}\t{甲[b.序-1].文本[:80]}"
                        for b, 源 in 中文待跟) + "\n", encoding="utf-8")
        print(f"✓ 待跟中文写进 {待跟路径.name}（{len(中文待跟)} 块）")
    print(f"✓ 已写回三方并更新快照")


def 存快照(甲, 乙, 条目id表, 反查):
    doc = Document(排版版)
    甲值, _ = md侧(甲, 反查)
    乙值, _ = md侧(乙, 反查)
    丙值, _, _ = docx侧(doc, 甲, 条目id表)
    块 = {}
    for b in 甲:
        if b.类型 not in 可同步:
            continue
        块[b.锚] = {"英文版": 甲值.get(b.序), "对照版": 乙值.get(b.序),
                   "Word": 丙值.get(b.序)}
    # 没有对照版时，它和英文版永远一致，不该参与「谁改了」的判断
    快照路径.write_text(json.dumps({"说明": "上次同步完成时三方的样子，别手删",
                                 "块": 块}, ensure_ascii=False, indent=0),
                     encoding="utf-8")


if __name__ == "__main__":
    主()
