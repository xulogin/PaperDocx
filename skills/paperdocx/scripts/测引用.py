# -*- coding: utf-8 -*-
r"""括号引用解析的自测：纯引用、说明文字 + 引用、不收的情况，以及 md 同步串来回不丢说明文字。

    py -3.12 -X utf8 scripts/测引用.py          # 全过打印「引用自测 N/N 通过」，退出码 0
"""
import re
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import 块模型 as M  # noqa: E402

渲染 = {"ref1": "Hallé et al., 1978", "ref2": "Nilson, 1971"}
反查 = {v: k for k, v in 渲染.items()}
反查["__正则__"] = M.引用正则("|".join(re.escape(s) for s in 渲染.values()))
例 = [
    ("纯引用", "see (Nilson, 1971) here", "see ⟦c:ref2⟧ here"),
    ("两条", "(Hallé et al., 1978; Nilson, 1971)", "⟦c:ref1;ref2⟧"),
    ("说明 + 引用", "x (architectural models; Hallé et al., 1978) y", "x ⟦c:ref1|architectural models⟧ y"),
    ("说明含分号", "(a; b; Nilson, 1971)", "⟦c:ref2|a; b⟧"),
    ("说明含括号", "(e.g., E(n)-equivariant networks; Nilson, 1971)", "⟦c:ref2|e.g., E(n)-equivariant networks⟧"),
    ("说明含公式不收", "($x$; Nilson, 1971)", None),
]
过 = 0
for 名, 入, 期 in 例:
    串, 公式, _ = (M.md转同步串(入, 反查) + (None,))[:3]
    if 期 is None:
        对 = "⟦c:" not in 串
    else:
        回 = M.同步串转md(串, 公式, 渲染)
        对 = 串 == 期 and 回 == 入
    print(("  ok  " if 对 else "  FAIL") + " " + 名 + "  " + 串)
    过 += 对
print(f"引用自测 {过}/{len(例)} 通过")
sys.exit(0 if 过 == len(例) else 1)
