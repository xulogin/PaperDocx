# -*- coding: utf-8 -*-
r"""交付前的结构体检：Zotero 字段、书目字段、ZOTERO_PREF_1、书签、样式。

    py -3.12 -X utf8 排版/验docx.py V49_排版版.docx

★ 这是**结构**检查，不是「看着对」。版面对不对只能真渲染出来看
  （渲染看图.ps1）。两样都做了才算验收过。
"""
from __future__ import annotations

import json
import re
import sys
import zipfile
from collections import Counter
from pathlib import Path

这里 = Path(__file__).resolve().parent
sys.path.insert(0, str(这里))
import _项目                                            # noqa: E402

根 = _项目.找根()
路 = Path(sys.argv[1]) if len(sys.argv) > 1 else None
if 路 is None:
    路 = sorted(根.glob("V*_排版版.docx"))[-1]
if not 路.is_absolute():
    路 = 根 / 路

z = zipfile.ZipFile(路)
doc = z.read("word/document.xml").decode("utf-8")
坏 = []

# 1. 行内引用字段
项 = re.findall(r"ADDIN ZOTERO_ITEM CSL_CITATION", doc)
书目 = re.findall(r"ADDIN ZOTERO_BIBL", doc)
print(f"文件：{路.name}  {路.stat().st_size/1e6:.1f} MB")
print(f"  行内引用字段 {len(项)} 处；书目字段 {len(书目)} 个")
if not 项:
    坏.append("一个行内引用字段都没有")
if len(书目) != 1:
    坏.append(f"书目字段应该正好 1 个，现在 {len(书目)} 个")

# 2. 每个字段的 uris 都要在 zot_match_index 里
索引 = json.loads((这里 / "引用数据" / "zot_match_index.json").read_text("utf-8"))
合法 = {v["uri"] for v in 索引.values()}
用到 = set(re.findall(r"http://zotero\.org/users/[^\"\\]+/items/\w+", doc))
野 = 用到 - 合法
print(f"  字段里用到 {len(用到)} 个条目 URI，其中 {len(用到 & 合法)} 个在索引里")
if 野:
    坏.append(f"有 {len(野)} 个 URI 不在索引里：{sorted(野)[:3]}")

# 3. Zotero 认领凭据
try:
    pref = z.read("docProps/custom.xml").decode("utf-8")
    print("  ZOTERO_PREF_1：" + ("在" if "ZOTERO_PREF_1" in pref else "⛔ 不在"))
    if "ZOTERO_PREF_1" not in pref:
        坏.append("docProps/custom.xml 里没有 ZOTERO_PREF_1，Zotero 不认这份文档")
    m = re.search(r'&lt;style id="([^"]+)"', pref) or re.search(r'<style id="([^"]+)"', pref)
    print(f"  样式 ID：{m.group(1) if m else '?'}")
except KeyError:
    坏.append("没有 docProps/custom.xml")

# 4. 书签（三方同步靠它）
import 排版配置 as C                                     # noqa: E402

书签 = re.findall(r'w:bookmarkStart[^>]*w:name="(%s\d+)"' % C.锚点前缀, doc)
print(f"  同步书签 {len(书签)} 个（{书签[0] if 书签 else '?'}…{书签[-1] if 书签 else '?'}）")
if len(书签) != len(set(书签)):
    坏.append("书签有重名")

# 5. 样式与字体
样式 = z.read("word/styles.xml").decode("utf-8")
主题字 = re.findall(r'(?:ascii|hAnsi|eastAsia|cs)Theme="[^"]+"', 样式)
print(f"  styles.xml 里残留的主题字属性 {len(主题字)} 处"
      + ("（★ 主题字优先于实名字体，标题会变无衬线）" if 主题字 else ""))
if 主题字:
    坏.append(f"styles.xml 里还有 {len(主题字)} 处主题字没换成实名")
用到样式 = Counter(re.findall(r'w:pStyle w:val="(\w+)"', doc))
print("  段落样式用量：", dict(用到样式.most_common(8)))

# 6. 参考文献条数
条 = doc.count('w:pStyle w:val="Bibliography"')
print(f"  Bibliography 段落 {条} 条")

print()
if 坏:
    print("⛔ 体检没过：")
    for b in 坏:
        print("   -", b)
    sys.exit(1)
print("✓ 结构体检通过（★ 版面还要真渲染出来肉眼过一遍：渲染看图.ps1）")
