# -*- coding: utf-8 -*-
r"""找项目根目录 + 找当前版本。**所有脚本都从这里拿路径，不许各自写死。**

项目根 = 往上找，第一个含有 `C_英文稿V*_英文版.md`（或 `稿件配置.py`）的目录。
所以 `排版/` 这个目录放在项目里的哪一层都行：

    你的项目/C_英文稿V42_英文版.md
    你的项目/排版/出Word.py                 ← 往上一层找到
    你的项目/_工作区/排版/出Word.py          ← 往上两层找到

版本 = 根目录里**字典序最大**的那个 `C_英文稿V*_英文版.md` 的版本号。

★ 2026-10-16 的教训：三方同步.py 把 V41 写死了，V42 开出来之后它还在同步
  一份已经冻结的旧稿，报「275 块」而当前稿是 266 块。**版本号一律动态取。**

## 换一套文件名怎么办

默认认 `C_英文稿V<数字>_英文版.md` / `_中英文对照版.md` / `_排版版.docx`。
要换名字，在项目根目录放一个 `稿件配置.py`：

    英文版名 = "manuscript_{版}_en.md"
    对照版名 = "manuscript_{版}_bilingual.md"
    排版版名 = "manuscript_{版}.docx"
    版正则   = r"manuscript_(v\d+)_en\.md"
"""
from __future__ import annotations

import re
import sys
from pathlib import Path

默认 = dict(
    英文版名="C_英文稿{版}_英文版.md",
    对照版名="C_英文稿{版}_中英文对照版.md",
    排版版名="C_英文稿{版}_排版版.docx",
    版正则=r"C_英文稿(V\d+)_英文版\.md",
    通配="C_英文稿V*_英文版.md",
)


def 找根(起=None) -> Path:
    r"""★ 先从**当前工作目录**往上找，再从脚本位置往上找。

    顺序不能反：skill 装好之后脚本在插件缓存里（`~/.claude/plugins/cache/...`），
    从那儿往上找永远找不到用户的项目。实测报
    「从 .../skills/paperdocx/scripts 往上找了一遍，都没有」。
    """
    试 = []
    for 起点 in ([Path(起).resolve()] if 起 else [Path.cwd(), Path(__file__).resolve()]):
        for d in [起点] + list(起点.parents):
            if not d.is_dir() or d in 试:
                continue
            试.append(d)
            if (d / "稿件配置.py").exists() or any(d.glob(默认["通配"])):
                return d
    raise SystemExit(
        "⛔ 找不到项目根目录。\n"
        "   判据是：某一层目录里有 `C_英文稿V*_英文版.md`，或者有 `稿件配置.py`。\n"
        f"   从当前目录 {Path.cwd()} 和脚本位置各往上找了一遍，都没有。\n"
        "   ⇒ **在项目目录里跑**，或者给 `--根 <路径>`。")


def 配置(根: Path) -> dict:
    c = dict(默认)
    f = 根 / "稿件配置.py"
    if f.exists():
        ns: dict = {}
        exec(compile(f.read_text("utf-8"), str(f), "exec"), ns)
        for k in 默认:
            if k in ns:
                c[k] = ns[k]
        if "通配" not in ns:
            c["通配"] = re.sub(r"\{版\}", "*", c["英文版名"])
    return c


def 当前版(根: Path, c: dict | None = None) -> str:
    c = c or 配置(根)
    稿 = sorted(根.glob(c["通配"]))
    if not 稿:
        raise SystemExit(f"⛔ {根} 里找不到任何英文版（通配 {c['通配']}）")
    m = re.search(c["版正则"], 稿[-1].name)
    if not m:
        raise SystemExit(f"⛔ {稿[-1].name} 对不上版正则 {c['版正则']}")
    return m.group(1)


def Zotero数据目录() -> Path:
    r"""Zotero 的数据目录。**只许有这一份解析逻辑。**

    优先级：`排版配置.文献["Zotero数据目录"]` > 环境变量 `ZOTERO_DATA_DIR` > `~/Zotero`。

    ★ 2026-10-19 实测教训：`做无DOI样式` 自己写了一套（只看环境变量和 `~/Zotero`），
      于是把改好的样式装进了 `C:\Users\<me>\Zotero\styles\` —— 那是 2022 年
      留下的**废弃 profile**，活的库在别的盘。装完一点报错都没有，
      作者在 Word 里 Refresh 还是用旧样式，现象是「改了等于没改」。
    """
    import os
    import sys
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    配 = None
    try:
        import 排版配置
        配 = 排版配置.文献.get("Zotero数据目录")
    except Exception:
        pass
    return Path(配 or os.environ.get("ZOTERO_DATA_DIR") or Path.home() / "Zotero")


def 找引用数据(根=None, 这里=None) -> Path:
    r"""`引用数据/`（references.json / zot_match_index.json / 引用映射.tsv / .csl）在哪。

    ★ 为什么要找好几处：skill 装好之后 `scripts/` 在插件缓存里（只读、所有项目共用），
      而引用数据是**每个项目各自的**，只能放项目里。写死 `这里/引用数据` 的话，
      直接跑 `<插件>/scripts/xxx.py` 就找不到。
    ★ 2026-10-19 实测：验docx.py 就是写死了 `这里/引用数据`，从插件安装位置跑
      demo 直接 FileNotFoundError。所以这个查找只许有**一份实现**，放这儿。
    """
    根 = Path(根) if 根 else 找根()
    这里 = Path(这里) if 这里 else Path(__file__).resolve().parent
    for d in (根 / "引用数据", 根 / "排版" / "引用数据",
              根 / "_工作区" / "排版" / "引用数据", 这里 / "引用数据"):
        if d.is_dir() and any(d.glob("*.json")):
            return d
    raise SystemExit(
        "⛔ 找不到 `引用数据/`（要有 references.json 和 zot_match_index.json）。\n"
        f"   找过：{根}/引用数据、{根}/排版/引用数据、{这里}/引用数据\n"
        "   ⇒ 先跑 `配文献.py --写` 生成，或见 references/zotero接法.md")


def 三件套(根=None, 版=None):
    """-> (根, 版, 英文版路径, 对照版路径, 排版版路径)"""
    根 = Path(根) if 根 else 找根()
    c = 配置(根)
    版 = 版 or 当前版(根, c)
    return (根, 版,
            根 / c["英文版名"].format(版=版),
            根 / c["对照版名"].format(版=版),
            根 / c["排版版名"].format(版=版))


if __name__ == "__main__":
    根, 版, a, b, d = 三件套(*sys.argv[1:])
    print(f"项目根：{根}\n当前版：{版}")
    for p in (a, b, d):
        print(f"  {'✓' if p.exists() else '✗'} {p.name}")
