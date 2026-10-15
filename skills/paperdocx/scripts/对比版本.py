# -*- coding: utf-8 -*-
r"""两个版本的对比体检 —— 回答「改完是不是真的比之前好」。

    py -3.12 -X utf8 _工作区/排版/对比版本.py                 # 自动取最新两版
    py -3.12 -X utf8 _工作区/排版/对比版本.py V41 V42

作者 2026-10-16：「改了以后要评估比之前好，几方面都比之前好才行哈。」
所以这里不只报词数 —— 词数变少太容易了，删掉内容就行。真正要证明的是
**「短了」的同时「一点东西都没少」，而且在别的维度上也没变差**。

## 六项

  ① 信息守恒 ★★ 最硬的一项：数字、引用、图表公式节的交叉引用、缩写，
     旧版有的新版必须还有。少一个就列出来，由人逐条确认是不是故意的。
  ② 词数：总量 + 分节，对照 WriteKit 实测的 RSE 分节基准
  ③ RSE 惯例：`按RSE惯例体检.py` 那张说法表，两版各数一遍，看有没有改坏
  ④ 可读性：平均句长、超长句（>45 词）条数、路标句密度
  ⑤ 术语一致性：point cloud / stem 那几条约定（和 `自检.py` ⑤ 同口径）
  ⑥ 结构：块数、图、表、公式、标题

★ 判据：① 必须「无丢失」或「丢失全部可解释」；②④ 要朝好的方向动；
  ③⑤⑥ 不许变差。有一项变差就回去改，别凑合。
"""
from __future__ import annotations

import re
import sys
from collections import Counter
from pathlib import Path

这里 = Path(__file__).resolve().parent
sys.path.insert(0, str(这里))
sys.path.insert(0, str(这里.parent))
import 块模型 as M                                    # noqa: E402

import _项目                                           # noqa: E402
根 = _项目.找根(这里)

# WriteKit 从 RSE 实测的分节词数（中位, 四分位）。见 `_工作区/按RSE惯例体检.py`
RSE基准 = {"Abstract": (336, (234, 404)), "1 Introduction": (1240, (1023, 1446)),
          "2 Materials and methods": (3067, (2600, 4296)),
          "3 Results": (1418, (845, 1816)), "4 Discussion": (2118, (1804, 2614)),
          "5 Conclusions": (302, (283, 308))}

说法 = [
    ("ablation（RSE 0/22 篇，该避）", r"ablation", "避"),
    ("it was found / shown（被动，RSE 1/22）", r"\bit (?:was|is) (?:found|shown|proposed)\b", "避"),
    ("remains poorly understood（四语料全 0）", r"remains? (?:poorly|not well|largely) understood", "避"),
    ("paves the way（全 0）", r"paves? the way", "避"),
    ("warrants further（全 0）", r"warrants? further", "避"),
    ("utilize/utilise（该用 use）", r"utiliz|utilis", "避"),
    ("certainly/undoubtedly/obviously", r"certainly|undoubtedly|obviously", "避"),
    ("however（RSE 22/22，该有）", r"\bhowever\b", "有"),
    ("compared with/to（21/22）", r"compared (?:with|to)", "有"),
    ("although/despite（让步，20/22）", r"\balthough\b|\bdespite\b", "有"),
    ("may/might（hedge，21/22）", r"\b(?:may|might)\b", "有"),
    ("likely（15/22）", r"\blikely\b", "有"),
    ("sensitivity analysis（9/22）", r"sensitivity analys[ie]s", "有"),
    ("limitation（18/22）", r"\blimitations?\b", "有"),
    ("can be attributed to（RSE 专用归因）", r"can be attributed to|attributable to", "看"),
    ("we + 实义动词（主动语态）", r"\bwe (?:propose|present|use|test|compare|report|find|show|"
                                r"evaluate|retrain|reconstruct|measure|combine|take|ask)\b", "看"),
]

路标 = re.compile(r"\b(?:in (?:two|three|four) respects|we report|is examined next|"
                  r"in this section|this section (?:reads|describes|presents)|"
                  r"the next section|we (?:first|then) (?:compare|report|describe)|"
                  r"in turn, answering)\b", re.I)


def 载(版本):
    p = _项目.三件套(根, 版本)[2]
    if not p.exists():
        raise SystemExit(f"⛔ 找不到 {p.name}")
    return M.读英文版(p)


def 词(s):
    return len(re.sub(r"!\[\]\([^)]*\)", " ", M.行内公式.sub(" X ", s)).split())


def 正文(块们):
    return "\n".join(b.文本 for b in 块们 if b.类型 in
                     ("para", "figcap", "tabcap", "tabnote", "keywords"))


def 分节(块们):
    出, 名 = {}, "（题名页）"
    for b in 块们:
        if b.类型 == "h2":
            名 = b.文本.lstrip("# ").strip()
        出[名] = 出.get(名, 0) + 词(b.文本)
    return 出


def 资产(块们):
    """这一版都「说了哪些具体东西」—— 丢了任何一个都是丢内容。"""
    全 = "\n".join(b.文本 for b in 块们)
    正 = 正文(块们)
    return {
        "数字": Counter(re.findall(r"(?<![\w.])[−+-]?\d+(?:\.\d+)?%?(?![\w])", 全)),
        "引用": Counter(re.findall(r"\(([^()]*?\b(?:19|20)\d{2}[a-z]?[^()]*?)\)", 正)),
        "图表公式节引用": Counter(re.findall(
            r"\b(?:Figure|Table|Section|Sections|Eq\.|Equation)\s+[0-9A-Z][\w.–\-]*", 全)),
        "缩写定义": Counter(re.findall(r"\(([A-Z]{2,}(?:-[A-Z]+)?)\)", 全)),
    }


def 可读(块们):
    ps = [b.文本 for b in 块们 if b.类型 == "para"]
    句 = [s for p in ps for s in re.split(r"(?<=[.?]) ", p.strip()) if s.strip()]
    长 = [词(s) for s in 句]
    return {
        "句数": len(句),
        "平均句长": round(sum(长) / max(1, len(长)), 1),
        "超长句 >45 词": sum(1 for x in 长 if x > 45),
        "超长句 >60 词": sum(1 for x in 长 if x > 60),
        "路标句": len(路标.findall("\n".join(ps))),
    }


def 术语(块们):
    t = re.sub(r"!\[\]\([^)]*\)", "", "\n".join(b.文本 for b in 块们))
    单用cloud = sum(1 for m in re.finditer(r"\b[Cc]louds?\b", t)
                  if not re.search(r"point[- ]$", t[max(0, m.start() - 7):m.start()], re.I))
    return {"单用 cloud（应为 point cloud）": 单用cloud,
            "trunk（应为 stem）": len(re.findall(r"\b[Tt]runks?\b", t))}


def 箭头(旧, 新, 越小越好=True):
    if 新 == 旧:
        return "＝"
    好 = (新 < 旧) if 越小越好 else (新 > 旧)
    return "↓ 好" if (好 and 越小越好) else "↑ 好" if 好 else ("↓ 差" if 新 < 旧 else "↑ 差")


def 主():
    版 = [a for a in sys.argv[1:] if re.fullmatch(r"V\d+", a)]
    if len(版) != 2:
        c = _项目.配置(根)
        有 = sorted(re.search(c["版正则"], p.name).group(1)
                    for p in 根.glob(c["通配"]))
        if len(有) < 2:
            raise SystemExit("⛔ 根目录至少要有两版英文版才能对比")
        版 = 有[-2:]
    旧名, 新名 = 版
    旧, 新 = 载(旧名), 载(新名)
    print(f"对比：{旧名} → {新名}\n" + "=" * 78)

    # ── ① 信息守恒
    print("\n① 信息守恒（旧版有、新版没有的东西）★★ 最硬的一项")
    a旧, a新 = 资产(旧), 资产(新)
    丢总 = 0
    for 类 in a旧:
        少 = a旧[类] - a新[类]
        丢总 += sum(少.values())
        if 少:
            print(f"   ⚠ {类}：少了 {sum(少.values())} 处")
            for k, v in list(少.items())[:14]:
                print(f"       ×{v}  {k}")
        else:
            print(f"   ✓ {类}：一个没少（{sum(a旧[类].values())} 处全在）")
    多 = {类: sum((a新[类] - a旧[类]).values()) for 类 in a旧}
    print(f"   （新增：{多}）")

    # ── ② 词数
    print("\n② 词数（括号是 WriteKit 实测的 RSE 中位 / 四分位）")
    f旧, f新 = 分节(旧), 分节(新)
    总旧, 总新 = sum(f旧.values()), sum(f新.values())
    for 名 in list(dict.fromkeys(list(f旧) + list(f新))):
        o, n = f旧.get(名, 0), f新.get(名, 0)
        基 = RSE基准.get(名)
        标 = f"  (RSE 中位 {基[0]}，四分位 {基[1][0]}–{基[1][1]})" if 基 else ""
        内 = ""
        if 基:
            内 = " ✓在四分位内" if 基[1][0] <= n <= 基[1][1] else " ⚠在四分位外"
        print(f"   {名[:34]:34s} {o:6d} → {n:6d}  ({n-o:+d}){标}{内}")
    print(f"   {'合计':34s} {总旧:6d} → {总新:6d}  ({总新-总旧:+d})")

    # ── ③ RSE 惯例
    print("\n③ RSE 惯例（『避』越少越好，『有』别掉到 0）")
    t旧, t新 = 正文(旧), 正文(新)
    坏 = 0
    for 名, 正则, 向 in 说法:
        o = len(re.findall(正则, t旧, re.I))
        n = len(re.findall(正则, t新, re.I))
        if o == n == 0:
            continue
        记 = 箭头(o, n, 越小越好=(向 == "避"))
        if 向 == "避" and n > o:
            坏 += 1
        if 向 == "有" and n == 0 and o > 0:
            坏 += 1
            记 = "⛔ 掉到 0"
        print(f"   [{向}] {名[:44]:44s} {o:4d} → {n:4d}  {记}")
    print(f"   变差的项：{坏}")

    # ── ④ 可读性
    print("\n④ 可读性")
    r旧, r新 = 可读(旧), 可读(新)
    for k in r旧:
        越小越好 = k != "句数"
        print(f"   {k:22s} {r旧[k]:8} → {r新[k]:8}   {箭头(r旧[k], r新[k], 越小越好)}")

    # ── ⑤ 术语
    print("\n⑤ 术语一致性（都该是 0）")
    y旧, y新 = 术语(旧), 术语(新)
    for k in y旧:
        print(f"   {k:34s} {y旧[k]:4d} → {y新[k]:4d}   {箭头(y旧[k], y新[k])}")

    # ── ⑥ 结构
    print("\n⑥ 结构（图表公式标题不该变）")
    c旧, c新 = Counter(b.类型 for b in 旧), Counter(b.类型 for b in 新)
    for k in sorted(set(c旧) | set(c新)):
        标 = "" if c旧[k] == c新[k] else ("  ← 变了" if k not in ("para",) else "  ← 段落数变了（删了过渡段）")
        print(f"   {k:10s} {c旧[k]:4d} → {c新[k]:4d}{标}")

    # ── 结论
    print("\n" + "=" * 78)
    行 = []
    行.append(("信息守恒", "无丢失" if 丢总 == 0 else f"少 {丢总} 处，逐条看上面 ①"))
    行.append(("词数", f"{总旧} → {总新}（{总新-总旧:+d}）"))
    行.append(("RSE 惯例", "没变差" if 坏 == 0 else f"⛔ {坏} 项变差"))
    行.append(("可读性", f"平均句长 {r旧['平均句长']} → {r新['平均句长']}；"
                      f"超长句 {r旧['超长句 >45 词']} → {r新['超长句 >45 词']}；"
                      f"路标句 {r旧['路标句']} → {r新['路标句']}"))
    行.append(("术语", "都是 0" if sum(y新.values()) == 0 else f"⛔ {y新}"))
    for k, v in 行:
        print(f"  {k:10s} {v}")


if __name__ == "__main__":
    主()
