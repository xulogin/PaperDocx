# -*- coding: utf-8 -*-
r"""稿件 `## References` 里的纯文本条目 ↔ Zotero 库，生成 Zotero 活引用要的三件套。

    py -3.12 -X utf8 排版/配文献.py            # 只报告，不写文件
    py -3.12 -X utf8 排版/配文献.py --写        # 写 排版/引用数据/*

前提：Zotero 开着（本地 API 在 :23119），稿子的 `## References` 一节是
Elsevier 风格的「姓, 名缩写., …, 年. 题名. 刊物 卷, 页.」纯文本。

产物：
    引用数据/references.json      CSL-JSON，id 改成 ref1..refN
    引用数据/zot_match_index.json {refN: {itemKey, itemID, uri}}
    引用映射.tsv                  refN ↔ 正文里的渲染形式（出Word.py 靠它定位引用）
    引用数据/未命中.md             库里找不到的，拿去 LitDoc 下载

★ 和 references/zotero接法.md 里「手动导出 CSL JSON」不同：这里直接走
  **Zotero 本地 API**（`http://127.0.0.1:23119/api/users/0/items?format=csljson`），
  不用人去点「导出条目」，而且 CSL-JSON 的 `id` 本身就是条目 URI，
  省掉一次按题名回查。itemID 还是只能从 zotero.sqlite 拿（API 不给）。

★ 本机坑：有 HTTP_PROXY 环境变量，urllib 连 127.0.0.1 也会走代理返回 502，
  看起来和「Zotero 没开」一模一样 —— 必须 ProxyHandler({})。
"""
from __future__ import annotations

import difflib
import json
import os
import re
import shutil
import sqlite3
import sys
import tempfile
import urllib.request
from pathlib import Path

这里 = Path(__file__).resolve().parent
sys.path.insert(0, str(这里))
import _项目                                            # noqa: E402
import 排版配置 as C                                     # noqa: E402

根 = _项目.找根()
英文版 = _项目.三件套(根)[2]
出目 = 这里 / "引用数据"
API = "http://127.0.0.1:23119/api/users/0"
开 = urllib.request.build_opener(urllib.request.ProxyHandler({}))


# ───────────────────────────────── Zotero 侧
def 库版本() -> str:
    r"""Zotero 本地库的版本号（`Last-Modified-Version` 响应头）。

    库里**任何**改动都会把它顶上去，所以拿它当缓存的失效判据，
    比「我觉得文献没改」靠谱 —— 文献是会一直改的。
    """
    r = 开.open(f"{API}/items/top?format=json&limit=1", timeout=30)
    return r.headers.get("Last-Modified-Version") or ""


def 取库条目(强制=False):
    r"""本地 API 分页拉全部顶层条目，每条带 CSL-JSON **和** itemKey。

    ★ 带缓存：一千条要拉十来次 HTTP，是这个脚本最慢的一段。
      缓存按**库版本号**失效（见 库版本()），库一改就自动重拉，
      没改就直接读 `引用数据/_库缓存.json`。`--刷新库` 可以强制重拉。

    ★★ 不能只拉 `format=csljson`：CSL-JSON 的 `id` **不一定是条目 URI**。
      条目的 Extra 里写了 `Citation Key: Lipman:2023hef`（Better BibTeX 的习惯）时，
      Zotero 就拿这个 citekey 当 `id`。实测本库 65 条里有 2 条是这样
      （Lipman:2023hef、Liu:2022yjp），从 `id` 反推 itemKey 会得到一个
      查不到 itemID 的假键，Refresh 时就命中不了条目。
      ⇒ 走 `format=json&include=csljson`，itemKey 从 `key` 字段拿，别从 id 猜。
    """
    缓存 = 出目 / "_库缓存.json"
    版 = 库版本()
    if not 强制 and 缓存.exists():
        try:
            d = json.loads(缓存.read_text("utf-8"))
            if d.get("库版本") == 版 and d.get("条目"):
                print(f"  库缓存命中（库版本 {版}，{len(d['条目'])} 条，没重拉）")
                return d["条目"]
        except (ValueError, KeyError):
            pass
    出, 起 = [], 0
    while True:
        u = f"{API}/items/top?format=json&include=csljson&limit=100&start={起}"
        批 = json.loads(开.open(u, timeout=60).read().decode("utf-8"))
        if not 批:
            break
        for r in 批:
            c = r.get("csljson") or {}
            # ★ include=csljson 给的是 JSON **字符串**，而且里面是**单元素数组**，
            #   不是对象。两层都要脱。
            if isinstance(c, str):
                c = json.loads(c)
            if isinstance(c, list):
                c = c[0] if c else {}
            c = dict(c)
            c["__key__"] = r["key"]
            出.append(c)
        起 += len(批)
        if len(批) < 100:
            break
    出目.mkdir(parents=True, exist_ok=True)
    缓存.write_text(json.dumps({"库版本": 版, "条目": 出}, ensure_ascii=False), "utf-8")
    print(f"  从 Zotero 拉了 {len(出)} 条顶层条目（库版本 {版}），已缓存")
    return 出


def 取itemID():
    """itemKey -> itemID。复制一份 sqlite 再读，绝不碰原库。

    数据目录：排版配置.文献['Zotero数据目录'] > 环境变量 ZOTERO_DATA_DIR > `~/Zotero`。
    ★ 很多人把库挪到别的盘了（本机在 F:\\Zotero_storage），别写死默认位置。
    """
    库 = Path(C.文献.get("Zotero数据目录")
             or os.environ.get("ZOTERO_DATA_DIR")
             or Path.home() / "Zotero") / "zotero.sqlite"
    if not 库.exists():
        raise SystemExit(f"⛔ 找不到 {库}；用 ZOTERO_DATA_DIR 指一下数据目录")
    临 = Path(tempfile.mkdtemp(prefix="zotsnap_")) / "zotero.sqlite"
    shutil.copy2(库, 临)
    con = sqlite3.connect(str(临))
    try:
        return {k: i for i, k in con.execute("SELECT itemID, key FROM items")}
    finally:
        con.close()
        shutil.rmtree(临.parent, ignore_errors=True)


# ───────────────────────────────── 稿件侧
def 读文献表():
    md = 英文版.read_text("utf-8")
    _, _, 尾 = md.partition("\n## References\n")
    出 = []
    for s in (l.strip() for l in 尾.split("\n")):
        if not s:
            continue
        m = re.search(r",\s*(\d{4})([a-z]?)\.\s", s)
        if not m:
            raise SystemExit(f"⛔ 认不出年份：{s[:90]}")
        作者串 = s[:m.start()]
        部分 = [p.strip() for p in 作者串.split(",") if p.strip()]
        出.append(dict(原=s, 年=m.group(1), 后缀=m.group(2),
                      姓=部分[0], 余=s[m.end():]))
    return 出


def 规(s: str) -> str:
    r"""题名规约。★ 别把非 ASCII 全删掉 —— 实测教训：

    `[^a-z0-9]+` 会把中文题名《基于图像序列的树木枝干L系统建模方法研究》压成 `"l"`，
    把《“.5 共用体”》压成 `"5"`，于是「`t in 余`」对任何含 l / 5 的英文题名都成立，
    一条中文笔记能和 LESS 那篇撞成同分，报「★歧义 1.00/1.00」。
    中日韩字符要留着，而且短题名不许走子串匹配（见 配() 里的 最短子串）。
    """
    s = s.lower().replace("‐", "-").replace("–", "-").replace("—", "-")
    s = re.sub(r"[^a-z0-9㐀-䶿一-鿿぀-ヿ]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


最短子串 = 14          # 规约后短于这个长度的题名，只许走整体相似度，不许走子串


def 配(条, 库, 库规):
    """返回 (命中条目, 说明)"""
    余 = 规(条["余"])
    年 = 条["年"]
    姓 = 规(条["姓"]).split()[-1] if 条["姓"] else ""
    候选 = []
    for it, t规 in zip(库, 库规):
        if not t规:
            continue
        年库 = str((it.get("issued", {}).get("date-parts") or [[""]])[0][0] or "")
        # 年份允许差 1：Elsevier 的「2026」常常是在线先发的 2025 条目
        if 年库 and abs(int(年库) - int(年)) > 1:
            continue
        if len(t规) < 最短子串:
            分 = difflib.SequenceMatcher(None, t规, 余[:len(t规) + 10]).ratio()
            if 分 < 0.90:
                continue
        elif 余.startswith(t规[:min(len(t规), 60)]):
            分 = 1.0
        elif t规 in 余:
            分 = 0.97
        else:
            分 = difflib.SequenceMatcher(None, t规[:70], 余[:70]).ratio()
            if 分 < 0.80:
                continue
        作者 = it.get("author") or it.get("editor") or []
        姓库 = 规(作者[0].get("family", "")) if 作者 else ""
        if 姓 and 姓库 and 姓 not in 姓库 and 姓库 not in 姓:
            分 -= 0.25
        候选.append((分, it))
    候选.sort(key=lambda x: -x[0])
    if not 候选:
        return None, "无候选"
    if len(候选) > 1 and 候选[0][0] - 候选[1][0] < 0.02:
        return 候选[0][1], f"★歧义 {候选[0][0]:.2f}/{候选[1][0]:.2f}"
    return 候选[0][1], f"{候选[0][0]:.2f}"


# ───────────────────────────────── 正文引用形式 ↔ refN
括号引用 = re.compile(r"\(([^()]*?\b(?:19|20)\d{2}[a-z]?)\)")
叙述引用 = re.compile(
    r"((?:[A-Z][\wÀ-ɏ'’\-]*)(?:,\s*[A-Z]\.)?"
    r"(?:\s+(?:and|&)\s+[A-Z][\wÀ-ɏ'’\-]*)?(?:\s+et al\.)?)"
    r"\s\(((?:19|20)\d{2}[a-z]?)\)")


def 收引用形式(正文: str):
    """-> (括号里的形式 set, 叙述式 [(作者串, 年)])"""
    形 = set()
    for m in 括号引用.finditer(正文):
        for p in m.group(1).split(";"):
            p = p.strip()
            # 只有年份的括号是叙述式引用的后半截（"Yang and Liu (2026)"），
            # 不是独立的括号引用，由 叙述引用 那条路管
            if re.search(r"[A-Za-z]", p):
                形.add(p)
    return 形, [(m.group(1), m.group(2)) for m in 叙述引用.finditer(正文)]


def 形式配条目(形式: str, 条们):
    """渲染形式 'Liu, C. et al., 2023' → 命中的文献条目下标列表（0 起）"""
    m = re.match(r"^(.*?),\s*((?:19|20)\d{2})([a-z]?)$", 形式)
    if not m:
        return []
    作, 年, 后缀 = m.group(1).strip().rstrip(","), m.group(2), m.group(3)
    作 = re.sub(r"\bet al\.?$", "", 作).strip().rstrip(",")
    初 = None
    m2 = re.match(r"^(.*?),\s*([A-Z]\.)$", 作)
    if m2:
        作, 初 = m2.group(1), m2.group(2)
    姓们 = [re.sub(r"^(?:de|van|von|der|den)\s+", "", s.strip())
           for s in re.split(r"\s+and\s+|\s+&\s+", 作)]
    出 = []
    for i, e in enumerate(条们):
        if e["年"] != 年 or e["后缀"] != 后缀:
            continue
        姓表 = [p.strip() for p in e["原"][:len(e["原"]) - len(e["余"])].split(",")]
        姓表 = [姓表[j] for j in range(0, len(姓表), 2)]
        if 姓们[0].lower() != 姓表[0].lower():
            continue
        if len(姓们) == 2 and (len(姓表) < 2 or 姓们[1].lower() != 姓表[1].lower()):
            continue
        if 初:
            缩 = [p.strip() for p in e["原"].split(",")]
            if len(缩) < 2 or not 缩[1].startswith(初[0]):
                continue
        出.append(i)
    return 出


def 建引用映射(条们):
    r"""扫正文，把每个渲染形式配到 refN，写 引用数据/引用映射.tsv。

    ★ 为什么不按 CSL 规则**生成**渲染形式：这份稿子里消歧是不规整的 ——
      `Liu, C. et al., 2023` 和 `Liu, X. et al., 2023` 带名缩写，
      同一个 Liu, C. 的 2022 那条却写成 `Liu et al., 2022`。
      按规则生成会和稿子里的字不一样，出 Word 时正则就对不上。
      ⇒ **以稿子里实际出现的形式为准**，反过来往条目上配。
    """
    正文 = 英文版.read_text("utf-8").partition("\n## References\n")[0]
    形, 叙述 = 收引用形式(正文)
    映, 坏 = {}, []
    for f in sorted(形):
        c = 形式配条目(f, 条们)
        if len(c) == 1:
            映[f"ref{c[0]+1}"] = f
        else:
            坏.append((f, len(c)))
    叙述未配 = []
    for 作, 年 in 叙述:
        f = f"{作}, {年}"
        c = 形式配条目(f, 条们)
        if len(c) == 1 and f"ref{c[0]+1}" not in 映:
            叙述未配.append((f"ref{c[0]+1}", f))
    return 映, 坏, 叙述, 叙述未配


# ───────────────────────────────── 主
def 主():
    写 = "--写" in sys.argv
    条们 = 读文献表()
    库 = 取库条目("--刷新库" in sys.argv)
    库规 = [规(it.get("title", "")) for it in 库]
    print(f"稿件文献 {len(条们)} 条；Zotero 顶层条目 {len(库)} 条")

    键表 = 取itemID()
    命中, 未中, 提醒 = {}, [], []
    for i, 条 in enumerate(条们, 1):
        it, 说 = 配(条, 库, 库规)
        标 = f"ref{i}"
        if it is None:
            未中.append((标, 条))
            print(f"  ✗ {标:6} {条['姓']} {条['年']}{条['后缀']}  {条['余'][:60]}")
            continue
        命中[标] = (条, it, 说)
        首分 = float((re.findall(r"\d\.\d+", 说) or ["0"])[0])
        if 说.startswith("★") or 首分 < 0.995:
            提醒.append((标, 条, it, 说))
    print(f"\n命中 {len(命中)} / 未命中 {len(未中)}")
    if 提醒:
        print("\n★ 这些命中不是逐字相同，人工看一眼：")
        for 标, 条, it, 说 in 提醒:
            print(f"  {标:6} [{说}]\n     稿：{条['姓']} {条['年']} | {条['余'][:70]}"
                  f"\n     库：{(it.get('author') or [{}])[0].get('family','?')}"
                  f" {((it.get('issued',{}).get('date-parts') or [['?']])[0][0])}"
                  f" | {it.get('title','')[:70]}")
    if 未中:
        print("\n⛔ 库里没有这些，要用 LitDoc 下载入库：")
        for 标, 条 in 未中:
            print(f"  {标}: {条['原'][:110]}")

    if not 写:
        print("\n（只报告，没写文件。确认无误后加 --写）")
        return

    出目.mkdir(parents=True, exist_ok=True)
    # URI 前缀从库里真实条目取（localUserKey 每台机器不一样，不许写死）
    前缀 = next((c["id"].rsplit("/", 2)[0] for c in 库
                if isinstance(c.get("id"), str) and c["id"].startswith("http")), None)
    if not 前缀:
        raise SystemExit("⛔ 库里一条 URI 形式的 id 都没有，取不到 URI 前缀")
    refs, 索引 = [], {}
    for 标, (条, it, 说) in 命中.items():
        c = dict(it)
        键 = c.pop("__key__")
        c["id"] = 标
        refs.append(c)
        索引[标] = dict(itemKey=键, itemID=键表.get(键, 0),
                      uri=f"{前缀}/items/{键}")
        if not 索引[标]["itemID"]:
            print(f"  ⚠ {标} 的 itemKey {键} 在 zotero.sqlite 里查不到 itemID")
    (出目 / "references.json").write_text(
        json.dumps(refs, ensure_ascii=False, indent=1), "utf-8")
    (出目 / "zot_match_index.json").write_text(
        json.dumps(索引, ensure_ascii=False, indent=1), "utf-8")
    if 未中:
        (出目 / "未命中.md").write_text(
            "# Zotero 库里找不到的文献（拿去 LitDoc 下载）\n\n"
            + "".join(f"{i+1}. {条['原']}\n" for i, (_, 条) in enumerate(未中)), "utf-8")
    print(f"\n✓ 写了 {出目/'references.json'}（{len(refs)} 条）"
          f"\n✓ 写了 {出目/'zot_match_index.json'}")
    # refN ↔ 条目原文，给 做排版源.py 对账用
    (出目 / "条目归属.json").write_text(json.dumps(
        {标: 命中[标][0]["原"] for 标 in 命中}, ensure_ascii=False, indent=1), "utf-8")

    # ---- 引用映射.tsv ----
    映, 坏, 叙述, 叙述未配 = 建引用映射(条们)
    if 坏:
        raise SystemExit("⛔ 这些正文引用形式配不上文献条目（或歧义），"
                         "先把稿子或文献表改对：\n  "
                         + "\n  ".join(f"{f!r} 命中 {n} 条" for f, n in 坏))
    行 = ["refN\t渲染形式"] + [f"{k}\t{v}" for k, v in
                           sorted(映.items(), key=lambda kv: int(kv[0][3:]))]
    (出目 / "引用映射.tsv").write_text("\n".join(行) + "\n", "utf-8")
    未引 = [标 for 标 in 命中 if 标 not in 映]
    print(f"✓ 写了 {出目/'引用映射.tsv'}（正文引用 {len(映)} 种）")
    if 未引:
        print(f"  ★ 正文没引、只在补充材料里引的 {len(未引)} 条："
              f"{','.join(未引)} → 要进参考文献表就写进 排版配置.文献['额外收录']")
    print(f"  ★ 叙述式引用 {len(叙述)} 处（出 Word 时要改写成 `作者 (作者, 年)`，"
          f"做排版源.py 负责）：" + "；".join(f"{a} ({y})" for a, y in 叙述))


if __name__ == "__main__":
    主()
