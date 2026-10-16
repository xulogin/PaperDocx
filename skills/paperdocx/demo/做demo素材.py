# -*- coding: utf-8 -*-
r"""给 demo 生成图和引用数据。**虚构内容，和任何真实研究无关。**

    py -3.12 -X utf8 做demo素材.py

生成：
    figures/Fig01_instrument.png   figures/Fig02_scatter.png
    引用映射.tsv
    引用数据/references.json
    引用数据/zot_match_index.json
    引用数据/demo-author-date.csl
"""
import json
import math
import random
from pathlib import Path

from PIL import Image, ImageDraw

这里 = Path(__file__).resolve().parent
random.seed(20240401)                                   # 固定种子，产物可复现


# ══════════════════════════════════════════ 图
def 画图():
    d = 这里 / "figures"
    d.mkdir(exist_ok=True)
    # 按 14 cm 宽 @300 dpi ≈ 1654 px 出，和正文里的 {width=14cm} 对得上
    W = 1654

    # Fig 1：三联图，示意仪器 / 现场 / 站点分布
    im = Image.new("RGB", (W, int(W * 0.42)), "white")
    g = ImageDraw.Draw(im)
    w3, h = W // 3, im.height
    for i, (标, 色) in enumerate([("(a) instrument", (70, 110, 160)),
                                 ("(b) deployment", (90, 140, 100)),
                                 ("(c) sites", (170, 120, 70))]):
        x0 = i * w3
        g.rectangle([x0 + 18, 18, x0 + w3 - 18, h - 18], outline=(120, 120, 120), width=3)
        g.text((x0 + 34, 34), 标, fill=(40, 40, 40))
    # (a) 膜 + 话筒
    cx, cy = w3 // 2, h // 2
    g.ellipse([cx - 150, cy - 44, cx + 150, cy + 44], outline=(70, 110, 160), width=6)
    g.ellipse([cx - 150, cy + 56, cx + 150, cy + 144], outline=(70, 110, 160), width=6)
    g.rectangle([cx - 26, cy + 150, cx + 26, cy + 210], fill=(70, 110, 160))
    for k in range(5):                                   # 落下的雨滴
        g.ellipse([cx - 110 + k * 54, cy - 160 - k % 3 * 26,
                   cx - 92 + k * 54, cy - 132 - k % 3 * 26], fill=(70, 110, 160))
    # (b) 两个并排的柱子
    bx = w3 + w3 // 2
    for off, col in ((-70, (90, 140, 100)), (70, (140, 140, 140))):
        g.rectangle([bx + off - 22, cy - 40, bx + off + 22, cy + 190], fill=col)
        g.ellipse([bx + off - 48, cy - 76, bx + off + 48, cy - 10], fill=col)
    g.line([bx - 260, cy + 192, bx + 260, cy + 192], fill=(80, 80, 80), width=5)
    # (c) 九个点，三组
    sx = 2 * w3 + w3 // 2
    for gi in range(3):
        for si in range(3):
            px = sx - 150 + gi * 150
            py = cy - 70 + si * 70
            g.ellipse([px - 16, py - 16, px + 16, py + 16],
                      fill=[(70, 110, 160), (120, 90, 150), (190, 130, 60)][gi])
    im.save(d / "Fig01_instrument.png", optimize=True)

    # Fig 2：三联散点图
    im = Image.new("RGB", (W, int(W * 0.40)), "white")
    g = ImageDraw.Draw(im)
    for i, (标, 噪, 饱) in enumerate([("Temperate oceanic", 0.08, None),
                                    ("Humid continental", 0.10, None),
                                    ("Tropical savanna", 0.12, 40)]):
        x0 = i * w3
        L, R = x0 + 110, x0 + w3 - 50
        T, B = 70, im.height - 86
        g.rectangle([L, T, R, B], outline=(120, 120, 120), width=3)
        g.line([L, B, R, T], fill=(150, 150, 150), width=3)          # 1:1 虚线位置
        g.text((L + 10, 34), 标, fill=(40, 40, 40))
        g.text((L + 10, B + 30), "reference (mm)", fill=(90, 90, 90))
        for _ in range(170):
            x = random.random() ** 1.7                                # 小值多
            y = x * (1 + random.gauss(0, 噪))
            if 饱 and x * 96 > 饱:                                     # 高强度处压低
                y *= 1 - 0.32 * min(1, (x * 96 - 饱) / 56)
            y = max(0.0, min(1.0, y))
            px, py = L + x * (R - L), B - y * (B - T)
            c = (200, 90, 70) if x > 0.6 else (70, 120, 170)
            g.ellipse([px - 5, py - 5, px + 5, py + 5], fill=c)
    im.save(d / "Fig02_scatter.png", optimize=True)
    print(f"✓ 图 2 张 → {d}")


# ══════════════════════════════════════════ 引用
文献 = [
    ("ref1", "Hollis et al., 2019",
     [("Hollis", "Daniel"), ("Arendt", "Lena"), ("Mbeki", "Sipho")],
     "Gauge density and the representativeness of urban flood forecasts",
     "Journal of Hydrometeorology", 2019, "20", "1455-1470"),
    ("ref2", "Okafor and Lindqvist, 2021",
     [("Okafor", "Chidi"), ("Lindqvist", "Erik")],
     "Low-cost precipitation sensing: a review of field deployments",
     "Environmental Monitoring and Assessment", 2021, "193", "612"),
    ("ref3", "Vasquez et al., 2018",
     [("Vasquez", "Rosa"), ("Ndiaye", "Awa"), ("Petrov", "Ilya")],
     "Impact acoustics of raindrops on tensioned membranes",
     "Measurement Science and Technology", 2018, "29", "085103"),
    ("ref4", "Chen and Møller, 2022",
     [("Chen", "Wei"), ("Møller", "Anders")],
     "Laboratory characterisation of an acoustic disdrometer",
     "Atmospheric Measurement Techniques", 2022, "15", "2201-2216"),
]


def 写引用():
    d = 这里 / "引用数据"
    d.mkdir(parents=True, exist_ok=True)

    csl = [{
        "id": rid, "type": "article-journal", "title": 题,
        "container-title": 刊, "volume": 卷, "page": 页,
        "issued": {"date-parts": [[年]]},
        "author": [{"family": f, "given": g} for f, g in 作],
        # ★ 真实项目里这里会有 DOI。不出 DOI 是**样式**的事，不是删数据。
        "DOI": f"10.0000/demo.{rid}",
    } for rid, _, 作, 题, 刊, 年, 卷, 页 in 文献]
    (d / "references.json").write_text(
        json.dumps(csl, ensure_ascii=False, indent=1), encoding="utf-8")

    # 真实项目里 itemKey / uri 要从 Zotero 的 zotero.sqlite 里读（见 references/zotero接法.md）
    索引 = {rid: {"itemKey": f"DEMOKEY{i:02d}", "itemID": 1000 + i,
                 "uri": f"http://zotero.org/users/local/demo/items/DEMOKEY{i:02d}"}
          for i, (rid, *_) in enumerate(文献, 1)}
    (d / "zot_match_index.json").write_text(
        json.dumps(索引, ensure_ascii=False, indent=1), encoding="utf-8")

    (这里 / "引用映射.tsv").write_text(
        "refN\t渲染引用\n" + "".join(f"{rid}\t{渲}\n" for rid, 渲, *_ in 文献),
        encoding="utf-8")

    (d / "demo-author-date.csl").write_text(CSL, encoding="utf-8")
    print(f"✓ 引用数据 {len(文献)} 条 → {d}")


# 一份最小可用的「著者—年份」CSL。真实投稿请换成刊物指定的那份
# （Zotero 数据目录的 styles/ 里有几千个）。
CSL = """<?xml version="1.0" encoding="utf-8"?>
<style xmlns="http://purl.org/net/xbiblio/csl" class="in-text" version="1.0"
       demote-non-dropping-particle="never" default-locale="en-GB">
  <info>
    <title>PaperDocx Demo (author-date)</title>
    <id>http://www.zotero.org/styles/paperdocx-demo</id>
    <link href="http://www.zotero.org/styles/paperdocx-demo" rel="self"/>
    <category citation-format="author-date"/>
    <updated>2024-04-01T00:00:00+00:00</updated>
  </info>
  <macro name="author">
    <names variable="author">
      <name and="text" delimiter-precedes-last="never" initialize-with="." name-as-sort-order="all"/>
    </names>
  </macro>
  <macro name="author-short">
    <names variable="author">
      <name form="short" and="text" delimiter-precedes-last="never"/>
      <et-al term="et-al"/>
    </names>
  </macro>
  <macro name="year"><date variable="issued"><date-part name="year"/></date></macro>
  <!-- ★ 不出 DOI 就是把这个宏掏空。Zotero 库里的 DOI 一个都不用删。 -->
  <macro name="access"><text value=""/></macro>
  <citation et-al-min="3" et-al-use-first="1" disambiguate-add-year-suffix="true"
            disambiguate-add-givenname="true" givenname-disambiguation-rule="all-names">
    <layout prefix="(" suffix=")" delimiter="; ">
      <group delimiter=", ">
        <text macro="author-short"/>
        <text macro="year"/>
      </group>
    </layout>
  </citation>
  <bibliography hanging-indent="true" et-al-min="8" et-al-use-first="7">
    <sort><key macro="author"/><key macro="year"/></sort>
    <layout suffix=".">
      <group delimiter=", ">
        <text macro="author"/>
        <text macro="year"/>
        <text variable="title"/>
        <text variable="container-title" font-style="italic"/>
        <text variable="volume"/>
        <text variable="page"/>
      </group>
      <text macro="access" prefix=" "/>
    </layout>
  </bibliography>
</style>
"""

if __name__ == "__main__":
    画图()
    写引用()
    print("\n接下来：")
    print("  py -3.12 -X utf8 ../scripts/出Word.py")
    print("  powershell -ExecutionPolicy Bypass -File ..\\scripts\\渲染看图.ps1")
