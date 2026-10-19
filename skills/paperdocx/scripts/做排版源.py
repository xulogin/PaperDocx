# -*- coding: utf-8 -*-
r"""英文版 md → **排版源** md。原件一个字节都不动。

    py -3.12 -X utf8 排版/做排版源.py            # 看它要改什么
    py -3.12 -X utf8 排版/做排版源.py --写        # 真写 <版>_排版源.md

为什么要这一步：`出Word.py` 要的 md 和作者手里的 md 有两处不一样。

1. **参考文献表**。作者的 md 里是 65 条纯文本；出 Word 要的是一个
   `(Reference list is inserted …)` 占位块 —— 文献表由 Zotero 书目字段生成，
   作者在 Word 里点 Refresh 才能重排、换样式。
2. ★ **叙述式引用**。`Yang and Liu (2026)` 这种写法，`出Word.py` 的引用正则
   只认 `(作者, 年)` 整个括号，**叙述式的括号里只有年份，匹配不上，
   出来就是一段死文本**，Refresh 不动它，换样式也不跟着变。
   改写成 `Yang and Liu (Yang and Liu, 2026)`，`插引用` 会识别出
   「前面已经写了作者」→ 给字段加 `suppress-author`，
   Word 里显示回 `Yang and Liu (2026)`，但它是活字段。

出来的 `<版>_排版源.md` 是**派生物**，别在上面改稿；改稿改 `<版>_英文版.md`，
再跑一遍这个脚本。
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

这里 = Path(__file__).resolve().parent
sys.path.insert(0, str(这里))
import _项目                                            # noqa: E402
import 配文献 as P                                       # noqa: E402

占位 = ("(Reference list is inserted here as a Zotero bibliography field; "
      "see 排版/引用数据/.)")


def 主():
    写 = "--写" in sys.argv
    根, 版, 英, _对, _排 = _项目.三件套()
    源 = 英.read_text("utf-8").replace("\r\n", "\n")
    出路 = 根 / f"{版}_排版源.md"

    # ---------- 1. 文献表换成占位块 ----------
    头, 分, 尾 = 源.partition("\n## References\n")
    if not 分:
        raise SystemExit("⛔ 找不到 `## References` 这一节")
    原条数 = len([l for l in 尾.split("\n") if l.strip()])
    新 = 头 + 分 + "\n" + 占位 + "\n"

    # ---------- 2. 叙述式引用改写 ----------
    条们 = P.读文献表()
    映 = {}
    引用数据 = _项目.找引用数据(根, 这里)
    for 行 in (引用数据 / "引用映射.tsv").read_text("utf-8").splitlines()[1:]:
        标, 形 = 行.split("\t")
        映[标] = 形
    改 = []

    def 换(m):
        作, 年 = m.group(1), m.group(2)
        c = P.形式配条目(f"{作}, {年}", 条们)
        if len(c) != 1:
            return m.group(0)
        标 = f"ref{c[0]+1}"
        形 = 映.get(标)
        if not 形:
            return m.group(0)
        改.append((m.group(0), f"{作} ({形})"))
        return f"{作} ({形})"

    新 = P.叙述引用.sub(换, 新)

    print(f"排版源：{出路.name}")
    print(f"  参考文献：{原条数} 行纯文本 → 1 个 Zotero 占位块")
    print(f"  叙述式引用改写 {len(改)} 处：")
    for 旧, 新文 in 改:
        print(f"    {旧}  →  {新文}")
    if not 写:
        print("\n（没写文件。确认无误后加 --写）")
        return
    出路.write_bytes(新.encode("utf-8"))
    print(f"\n✓ 写了 {出路}（{len(新)} 字符）")
    print(f"  ★ 原件 {英.name} 未改动")


if __name__ == "__main__":
    主()
