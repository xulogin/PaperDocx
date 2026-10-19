# -*- coding: utf-8 -*-
r"""从 md 到验收过的 docx，**一条命令走完**。

    py -3.12 -X utf8 排版/走一遍.py                 # 预览版（图 300 dpi）+ 全部验收
    py -3.12 -X utf8 排版/走一遍.py --交付           # 预览验收过了，再出一版原图交付稿
    py -3.12 -X utf8 排版/走一遍.py --版 V39         # 目录里堆了好几版时，指定排哪一版
    py -3.12 -X utf8 排版/走一遍.py --出目录 new      # 产物放进 <项目>/new/，不堆在根目录
    py -3.12 -X utf8 排版/走一遍.py --跳过文献        # 只改了正文、文献一个没动时
    py -3.12 -X utf8 排版/走一遍.py --刷新库          # 强制重拉 Zotero 库（平时靠库版本号自动失效）

## 为什么是这个顺序

    配文献 ──→ 做排版源 ──→ 出Word ──→ 渲染看图 ──→ 比文献表
       │                                              ↑
       └──────────→ 核DOI（联网，和上面这条路无关）───────┘

**前三步是硬链**：出 Word 要 `引用映射.tsv` 才知道正文里哪一段是引用。
**核DOI 是另一条路**，65 条要联网问 65 次，和出 docx 互不相干 ——
所以它一开始就**扔到后台**跑，最后再收结果。

## ★ 慢在哪（2026-10-18 第一次做这篇 65 条的稿子，花了一个多小时）

| 慢的 | 真正的原因 | 现在怎么解决 |
|---|---|---|
| 拉 Zotero 库 | 998 条要分 10 次 HTTP | 按**库版本号**缓存，库没改就不重拉 |
| 出了 4 遍 Word | 每遍都写几十 MB | 默认只出 **300 dpi 预览版**，验收过了再 `--交付` |
| DOI 核实 65 次 | 串在主路上等 | **扔后台**，和出 docx 同时跑 |
| 摸 Zotero API 格式来回 3 轮 | 照着印象写代码 | 已经写死在 配文献.py 的注释里，不用再摸 |

★ **并行不是万能的。** 这里真正省时间的是「库缓存」和「少出一遍 Word」，
  并行只省下 DOI 那一段。别为了并行把硬链拆开。
"""
from __future__ import annotations

import subprocess
import sys
import time
from pathlib import Path

这里 = Path(__file__).resolve().parent
sys.path.insert(0, str(这里))
import _项目                                            # noqa: E402
import 块模型 as M                                       # noqa: E402

def 取参(名, 缺省=None):
    return sys.argv[sys.argv.index(名) + 1] if 名 in sys.argv else 缺省


根 = _项目.找根()
# ★ `--版 V39`：一个项目目录里常常堆着 V39…V42 好几版，`三件套` 默认取字典序
#   最大的那一版。要排旧版就显式指定，别去动目录。
_, 版, 英, _对, 排 = _项目.三件套(根, 取参("--版"))
# ★ `--出目录 new`：产物（预览 / 交付稿）放进项目下的某个子目录，不堆在根目录。
出目录 = 取参("--出目录", "")
if 出目录:
    (根 / 出目录).mkdir(parents=True, exist_ok=True)
PY = [sys.executable, "-X", "utf8"]
交付 = "--交付" in sys.argv


def 出路(名):
    return str(Path(出目录) / 名) if 出目录 else 名

# ★ 两种 md，走的路不一样（2026-10-19 从插件安装位置跑 demo 才发现：
#   原来一律先跑 配文献.py，撞上 demo 的占位块就报「认不出年份」）：
#
#   原生格式：文献表已经是 `(Reference list is inserted …)` 占位块，
#            引用数据/ 也已经备好（demo、以及上一轮已经配过的稿子）
#            → 配文献 / 做排版源 都不用跑，md 自己就是排版源
#
#   作者手写：`## References` 底下 65 条纯文本
#            → 先配文献，再派生出 <版>_排版源.md
原生 = any(b.类型 == "bibph" for b in M.读英文版(英))


def 跑(名, *参, 必须=True):
    t0 = time.time()
    print(f"\n{'='*66}\n▶ {名}  {' '.join(str(x) for x in 参)}\n{'='*66}")
    r = subprocess.run(PY + [str(这里 / 名)] + [str(x) for x in 参], cwd=str(根))
    print(f"— {名} 用时 {time.time()-t0:.0f} 秒")
    if r.returncode and 必须:
        raise SystemExit(f"⛔ {名} 没过，停在这里。")
    return r.returncode


def 主():
    总 = time.time()

    # ── 0. 后台：DOI 核实（联网，和主路无关）。文献数据还没生成时跳过
    后台 = None
    if (这里 / "引用数据" / "references.json").exists() and "--跳过DOI" not in sys.argv:
        print("▶ 后台起 核DOI.py（和出 Word 同时跑）")
        后台 = subprocess.Popen(PY + [str(这里 / "核DOI.py")], cwd=str(根),
                              stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                              encoding="utf-8", errors="replace")

    # ── 1-3. 硬链
    if 原生:
        print(f"\n▶ {英.name} 已经是 PaperDocx 原生格式（文献表是 Zotero 占位块），"
              f"\n  配文献 / 做排版源 都跳过，直接拿它出 Word。")
        源 = 英.name
    else:
        if "--跳过文献" not in sys.argv:
            跑("配文献.py", "--写", *(["--刷新库"] if "--刷新库" in sys.argv else []))
        跑("做排版源.py", "--写")
        源 = f"{版}_排版源.md"
    预览 = 出路(f"_预览_{版}.docx")
    跑("出Word.py", "--源", 源, "--图dpi", "300", "--出", 预览)

    # ── 4. 真 Word 导 PDF → PNG
    t0 = time.time()
    # ★ 预览图落在**项目自己的** <根>/<出目录>/_预览，不落在插件脚本目录 ——
    #   脚本目录是所有项目共用的，两个项目的 p001.png 会互相覆盖。
    预览目录 = 根 / 出目录 / "_预览" if 出目录 else 根 / "_预览"
    print(f"\n{'='*66}\n▶ 渲染看图.ps1  {预览}  →  {预览目录}\n{'='*66}")
    r = subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File",
                        str(这里 / "渲染看图.ps1"), "-Docx", 预览,
                        "-Out", str(预览目录)], cwd=str(根))
    print(f"— 渲染 用时 {time.time()-t0:.0f} 秒")
    if r.returncode:
        raise SystemExit("⛔ 渲染没过，停在这里。")
    pdf = 预览目录 / Path(预览).name.replace(".docx", ".pdf")

    # ── 5. 验收
    跑("验docx.py", 预览)
    if not 原生:          # 原生格式的 md 里没有纯文本文献表，没得比
        跑("比文献表.py", pdf, 必须=False)

    # ── 6. 收后台
    if 后台:
        print(f"\n{'='*66}\n▶ 核DOI.py（后台）结果\n{'='*66}")
        out, _ = 后台.communicate()
        for l in (out or "").splitlines():
            if l.startswith(("  ✗", "  ⚠", "DOI 核实", "★")) or "没有 DOI" in l:
                print(l)

    print(f"\n{'='*66}")
    if 交付:
        交 = 出路(排.name)
        跑("出Word.py", "--源", 源, "--出", 交)
        print(f"✓ 交付稿 {交}（原图未压缩）")
    else:
        print(f"✓ 预览版 {预览} 已出并验收。版面没问题再跑一次带 --交付 出原图稿。")
    print(f"★ 一页页看图：{预览目录}")
    print(f"总用时 {time.time()-总:.0f} 秒")


if __name__ == "__main__":
    主()
