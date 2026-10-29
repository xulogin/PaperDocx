# -*- coding: utf-8 -*-
r"""后处理.py 的自测：拿 demo 搭一个临时项目，加进三种表 + 标红，出 Word、后处理，逐项核。

    py -3.12 -X utf8 scripts/测后处理.py          # 全过打印「后处理自测 N/N 通过」，退出码 0

核的东西：
  1. 两行表头：组名横向合并（gridSpan）、单列表头竖向合并（vMerge），两行都是表头样式、跨页重复
  2. 分组表：首列留空的格竖向合并
  3. 列宽：数值格「0.185 ± 0.070」用不换行空格连住、列宽放得下整串；总宽不变；字号 9 pt
  4. 标红：⟦…⟧ 的括号去掉、里面的字是 C00000
不碰 demo 目录本身（复制到临时目录做）。
"""
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

这里 = Path(__file__).resolve().parent
demo = 这里.parent / "demo"
PY = [sys.executable, "-X", "utf8"]

加的 = r"""

**Table 9.** Two-row header test

| Variable | Site A |  |  | Site B |  |  |
|---|--:|--:|--:|--:|--:|--:|
|  | Mean | SD | Bias | Mean | SD | Bias |
| Rain rate (mm h⁻¹) | 0.185 ± 0.070 | 1.2 | +3.0% | 0.176 ± 0.135 | 1.4 | −2.2% |
| Allometric exponent | 0.32 [0.18, 0.44] | 0.1 | Overlap | 0.34 [0.26, 0.40] | 0.1 | Overlap |

**Table 10.** Grouped first column test

| Stand | Scene | Value |
|---|---|--:|
| Larch | A | 1.0 |
|  | B | 2.0 |
|  | C | 3.0 |
| Birch | A | 4.0 |
|  | B | 5.0 |

The corresponding e-mail address is ⟦[to be completed]⟧ for now.
"""


def main():
    过, 共 = 0, 0

    def 核(条件, 话):
        nonlocal 过, 共
        共 += 1
        过 += bool(条件)
        print(("  ✓ " if 条件 else "  ✗ ") + 话)

    with tempfile.TemporaryDirectory() as td:
        项目 = Path(td) / "demo"
        shutil.copytree(demo, 项目, ignore=shutil.ignore_patterns("_预览*", "*.docx"))
        md = 项目 / "C_英文稿V01_英文版.md"
        s = md.read_text("utf-8")
        s = s.replace("\n## References", 加的 + "\n## References", 1)
        md.write_text(s, "utf-8")
        出 = 项目 / "测.docx"
        r = subprocess.run(PY + [str(这里 / "出Word.py"), "--图dpi", "100", "--出", str(出)],
                           cwd=项目, capture_output=True, text=True, encoding="utf-8", errors="replace")
        if r.returncode or not 出.exists():
            print(r.stdout[-1500:], r.stderr[-1500:])
            sys.exit("⛔ 出Word.py 没出来，后处理没法测")
        r = subprocess.run(PY + [str(这里 / "后处理.py"), str(出), str(出)],
                           cwd=项目, capture_output=True, text=True, encoding="utf-8", errors="replace")
        print(r.stdout.strip())
        核(r.returncode == 0, "后处理.py 跑完没报错")

        from docx import Document
        from docx.oxml.ns import qn
        d = Document(str(出))

        def 找表(首格):
            for t in d.tables:
                if t.rows[0].cells[0].text.strip() == 首格 and len(t.rows) >= 3:
                    return t
        t9 = 找表("Variable")
        核(t9 is not None, "找到两行表头测试表")
        if t9 is not None:
            tr0 = t9.rows[0]._tr
            spans = [int(tc.tcPr.find(qn("w:gridSpan")).get(qn("w:val")))
                     for tc in tr0.tc_lst if tc.tcPr is not None and tc.tcPr.find(qn("w:gridSpan")) is not None]
            核(spans == [3, 3], f"组名横向合并成两组各 3 列（实得 {spans}）")
            vm = tr0.tc_lst[0].tcPr.find(qn("w:vMerge"))
            核(vm is not None, "首列表头上下合并（vMerge）")
            核(t9.rows[1]._tr.trPr is not None and t9.rows[1]._tr.trPr.find(qn("w:tblHeader")) is not None,
               "第二行表头跨页重复")
            格 = [c.text for c in t9.rows[2].cells]
            核(any(" ± " in x for x in 格), "「±」两边是不换行空格")
            cols = [int(c.get(qn("w:w"))) / 567 for c in t9._tbl.find(qn("w:tblGrid")).findall(qn("w:gridCol"))]
            import importlib
            sys.path.insert(0, str(这里))
            H = importlib.import_module("后处理")
            需 = H._字宽cm("0.185 ± 0.070", 9.0) + H._格边距cm(t9)
            核(cols[1] >= 需 - 0.01, f"数值列放得下整串（列宽 {cols[1]:.2f} ≥ 需要 {需:.2f} cm）")
            szs = {int(x.get(qn("w:val"))) for x in t9._tbl.iter(qn("w:sz"))}
            核(not szs or szs <= {18}, f"字号 9 pt（实得 {sorted(szs)}）")
        t10 = 找表("Stand")
        核(t10 is not None, "找到分组表测试表")
        if t10 is not None:
            vms = [tc.tcPr.find(qn("w:vMerge")) is not None for row in t10.rows[1:]
                   for tc in row._tr.tc_lst[:1] if tc.tcPr is not None]
            核(sum(vms) >= 4, f"首列空格竖向合并（vMerge 格数 {sum(vms)}）")
        红 = [r for p in d.paragraphs for r in p.runs
              if r.font.color is not None and r.font.color.rgb is not None and str(r.font.color.rgb) == "C00000"]
        核(any("to be completed" in r.text for r in 红), "⟦…⟧ 里的字标成红色")
        核(not any("⟦" in p.text or "⟧" in p.text for p in d.paragraphs), "⟦⟧ 括号已去掉")

    print(f"\n后处理自测 {过}/{共} 通过")
    sys.exit(0 if 过 == 共 else 1)


if __name__ == "__main__":
    main()
