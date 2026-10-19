# -*- coding: utf-8 -*-
r"""拿 references.json 的 DOI 去 Crossref 核实题名 / 年份 / 刊物。**不许编文献。**

    py -3.12 -X utf8 排版/核DOI.py

没有 DOI 的条目（书、报告、会议论文）只列出来，不算错。
★ 本机有 HTTP_PROXY，走代理才能出网；Crossref 要 UA 带邮箱才不被限流。
"""
from __future__ import annotations

import difflib
import json
import re
import sys
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

这里 = Path(__file__).resolve().parent
sys.path.insert(0, str(这里))
import _项目                                            # noqa: E402

引用数据 = _项目.找引用数据(这里=这里)
引 = json.loads((引用数据 / "references.json").read_text("utf-8"))
归表 = 引用数据 / "条目归属.json"
归 = json.loads(归表.read_text("utf-8")) if 归表.exists() else {}
UA = "PaperDocx/1.0 (mailto:hifengxu@gmail.com)"


def 规(s):
    s = re.sub(r"<[^>]+>", " ", s or "").lower()
    return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9]+", " ", s)).strip()


def 查(doi):
    r"""先问 Crossref；404 就退到 doi.org 内容协商。

    ★ 实测教训：**Crossref 404 不等于 DOI 是编的。** 这 5 条就不在 Crossref 名下 ——
      CNKI（10.13287 / 10.13466）、Eurographics（10.2312）、J-Stage（10.18960）、
      arXiv（10.48550）分别挂在 JaLC / DataCite 下。
      `https://doi.org/<doi>` 带 `Accept: application/vnd.citationstyles.csl+json`
      能一视同仁地把它们取回来，CSL-JSON 字段名和 Crossref 的一样。
    """
    for u, h in (("https://api.crossref.org/works/" + urllib.parse.quote(doi, safe=""),
                  {"User-Agent": UA}),
                 ("https://doi.org/" + urllib.parse.quote(doi, safe="/"),
                  {"User-Agent": UA,
                   "Accept": "application/vnd.citationstyles.csl+json"})):
        try:
            d = json.loads(urllib.request.urlopen(
                urllib.request.Request(u, headers=h), timeout=30).read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code != 404:
                raise
            continue
        return d["message"] if "message" in d else d
    raise urllib.error.HTTPError(u, 404, "两处都查不到", None, None)


无doi, 对, 疑, 错 = [], 0, [], []
for c in 引:
    doi = c.get("DOI")
    标 = c["id"]
    if not doi:
        无doi.append((标, c.get("type", "?"), (c.get("title") or "")[:60]))
        continue
    try:
        m = 查(doi)
    except urllib.error.HTTPError as e:
        错.append((标, doi, f"HTTP {e.code}"))
        print(f"  ✗ {标} {doi} → HTTP {e.code}")
        continue
    except Exception as e:
        错.append((标, doi, repr(e)))
        print(f"  ✗ {标} {doi} → {e!r}")
        continue
    题库 = 规(c.get("title"))
    # ★ Crossref 的 title 是**数组**，doi.org 内容协商给的是**字符串**。
    #   不分情况写 `[0]` 的话，字符串会被切成第一个字符。
    t网 = m.get("title") or ""
    题网 = 规(t网[0] if isinstance(t网, list) else t网)
    比 = difflib.SequenceMatcher(None, 题库, 题网).ratio()
    年库 = str((c.get("issued", {}).get("date-parts") or [[""]])[0][0] or "")
    年网 = str((m.get("issued", {}).get("date-parts") or [[""]])[0][0] or "")
    if 比 > 0.92 and (not 年库 or not 年网 or abs(int(年库) - int(年网)) <= 1):
        对 += 1
        print(f"  ✓ {标:6} {doi}")
    else:
        疑.append((标, doi, 比, 年库, 年网, c.get("title"),
                  t网[0] if isinstance(t网, list) else t网))
        print(f"  ⚠ {标:6} {doi}  题名相似度 {比:.2f}  年 {年库}/{年网}")

print(f"\nDOI 核实：对上 {对}，可疑 {len(疑)}，查不到 {len(错)}，无 DOI {len(无doi)}")
for 标, doi, 比, 年库, 年网, t1, t2 in 疑:
    print(f"\n★ {标} {doi}  {比:.2f}  {年库}/{年网}\n   库：{t1}\n   网：{t2}"
          f"\n   稿：{归.get(标, '')[:100]}")
if 无doi:
    print("\n没有 DOI 的（书 / 报告 / 会议论文，正常）：")
    for 标, t, ti in 无doi:
        print(f"   {标:6} {t:22} {ti}")
sys.exit(1 if 错 else 0)
