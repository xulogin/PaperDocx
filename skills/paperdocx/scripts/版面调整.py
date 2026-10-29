# -*- coding: utf-8 -*-
r"""图表往后挪一段，让页面尽量填满（Word 不会自己浮动图表，放不下就整块推到下一页、留下大片空白）。

    py -3.12 -X utf8 排版/版面调整.py                # 预览出版、量空白、挪、再出，最多 15 轮
    py -3.12 -X utf8 排版/版面调整.py --轮 5 --阈 0.15

每轮：
  做排版源 → 出Word（300 dpi 预览）→ 后处理 → Word 导 PDF → 量每页底部空白
  → 找第一个「本页空了 ≥ 阈、下一页开头是某张图/表」的地方
  → 在 <版>_英文版.md 里把那张图（图片块 + 图注块）或表（表题 + 表 + 表注）挪到它后面那段正文之后
  → 下一轮。挪不动（后面紧跟一级标题、或已经挪过两次）就记下跳过。

只换图表位置，不改任何文字。挪前备份 md 到 _预览/版面调整前_<时间>.md，每一步写进 _预览/版面调整.log。
★ 不跨 `## ` 一级章节（图跑进别的章节读者找不到）；二、三级小节可以跨。
"""
import re
import shutil
import subprocess
import sys
import time
from pathlib import Path

这里 = Path(__file__).resolve().parent
sys.path.insert(0, str(这里))
import _项目  # noqa: E402

根 = _项目.找根()
_, 版, 英, _对, _排 = _项目.三件套(根, None)
PY = [sys.executable, "-X", "utf8"]
预览 = 根 / "_预览"
预览.mkdir(exist_ok=True)
日志 = 预览 / "版面调整.log"


def 取参(名, 缺省):
    return type(缺省)(sys.argv[sys.argv.index(名) + 1]) if 名 in sys.argv else 缺省


轮数, 阈 = 取参("--轮", 15), 取参("--阈", 0.15)


def 记(话):
    print(话)
    with open(日志, "a", encoding="utf-8") as f:
        f.write(time.strftime("%H:%M:%S ") + 话 + "\n")


def 出版():
    """出一版 PDF，返回 PDF 路径。"""
    docx = 预览 / "_版面.docx"
    subprocess.run(PY + [str(这里 / "做排版源.py"), "--写"], cwd=根, capture_output=True)
    源 = f"{版}_排版源.md" if (根 / f"{版}_排版源.md").exists() else 英.name
    r = subprocess.run(PY + [str(这里 / "出Word.py"), "--源", 源, "--图dpi", "200", "--出", str(docx)],
                       cwd=根, capture_output=True, text=True, encoding="utf-8", errors="replace")
    if r.returncode:
        sys.exit("出Word 失败：" + r.stdout[-800:] + r.stderr[-800:])
    subprocess.run(PY + [str(这里 / "后处理.py"), str(docx), str(docx)], cwd=根, capture_output=True)
    out = 预览 / "_版面"
    subprocess.run(["powershell", "-ExecutionPolicy", "Bypass", "-File", str(这里 / "渲染看图.ps1"),
                    "-Docx", str(docx), "-Out", str(out), "-Dpi", "40"], cwd=根, capture_output=True)
    return out / "_版面.pdf"


def 量(pdf):
    """-> [(页号, 空白比例, 下一页开头的图/表名, 本页剩余高 pt, 那张图高 pt, 图注高 pt)]"""
    import fitz
    d = fitz.open(pdf)
    H = d[0].rect.height
    结果 = []
    for i, p in enumerate(d):
        底 = 0
        for b in p.get_text("blocks"):
            if b[3] < H - 60:                       # 页脚页码不算
                底 = max(底, b[3])
        for img in p.get_images():
            for r in p.get_image_rects(img[0]):
                底 = max(底, r.y1)
        剩 = H - 70 - 底
        空 = 剩 / (H - 140)
        名, 图高, 注高 = None, 0.0, 0.0
        if i + 1 < len(d):
            q = d[i + 1]
            头 = re.sub(r"^\s*\d+\s*", "", q.get_text()).strip()[:40]
            m = re.match(r"(Figure|Table) (S?\d+)\.", 头)
            if m:
                名 = f"{m.group(1)} {m.group(2)}"
                if m.group(1) == "Figure":
                    rs = [r for img in q.get_images() for r in q.get_image_rects(img[0])]
                    if rs:
                        top = min(rs, key=lambda r: r.y0)
                        图高 = top.height
                        for b in q.get_text("blocks"):
                            if b[4].lstrip().startswith(f"Figure {m.group(2)}."):
                                注高 = b[3] - top.y1     # 图底到图注底（含间距）
        结果.append((i + 1, 空, 名, 剩, 图高, 注高))
    return 结果


最小比例 = 0.88       # 图最多缩到原来的 88%，再小就宁可往后挪一段


def 缩(名, 比例):
    """把 md 里这张图的 {width=Xcm} 乘以比例（留 0.2 cm 余量）。返回 (旧, 新) 或 None。"""
    md = 英.read_text("utf-8")
    块 = 分块(md)
    位 = 单元(块, 名)
    if not 位:
        return None
    起, 止 = 位
    图块 = 块[起]
    m = re.search(r"\{width=([\d.]+)cm\}", 图块)
    if not m or not 图块.lstrip().startswith("!["):
        return None
    旧 = float(m.group(1))
    新 = max(round(旧 * 比例 - 0.2, 1), round(旧 * 最小比例, 1))
    块[起] = 图块[:m.start(1)] + f"{新:g}" + 图块[m.end(1):]
    英.write_text("\n\n".join(块) + "\n", "utf-8")
    return 旧, 新


def 分块(md):
    return re.split(r"\n{2,}", md.strip("\n"))


def 单元(块, 名):
    """名 = 'Figure 9' / 'Table 5' → (起, 止)（含）"""
    类, 号 = 名.split()
    for i, b in enumerate(块):
        if b.startswith(f"**{类} {号}.**"):
            if 类 == "Figure":                     # 图片块在图注前面
                起 = i - 1 if i > 0 and 块[i - 1].lstrip().startswith("![") else i
                return 起, i
            止 = i
            if i + 1 < len(块) and 块[i + 1].lstrip().startswith("|"):
                止 = i + 1
                if 止 + 1 < len(块) and 块[止 + 1].startswith("Note:"):
                    止 += 1
            return i, 止
    return None


def 是正文段(b):
    b = b.lstrip()
    return not (b.startswith(("#", "!", "|", "$$", "<!--", "**Figure", "**Table", "Note:")))


def 挪(名, 已挪):
    md = 英.read_text("utf-8")
    块 = 分块(md)
    位 = 单元(块, 名)
    if not 位:
        return f"找不到 {名}"
    起, 止 = 位
    k = 止 + 1
    while k < len(块) and not 是正文段(块[k]):
        if 块[k].startswith("## "):
            return f"{名} 后面是一级标题，不挪"
        if 块[k].lstrip().startswith(("![", "**Figure", "**Table")):
            return f"{名} 后面紧跟别的图表，不挪"
        k += 1
    if k >= len(块) or 块[k].startswith("## ") or "## References" in 块[k] or 块[k].startswith("## CRediT"):
        return f"{名} 后面没有正文段可借"
    新 = 块[:起] + 块[止 + 1:k + 1] + 块[起:止 + 1] + 块[k + 1:]
    英.write_text("\n\n".join(新) + "\n", "utf-8")
    已挪[名] = 已挪.get(名, 0) + 1
    return None


def main():
    备份 = 预览 / f"版面调整前_{time.strftime('%m%d_%H%M%S')}.md"
    shutil.copy(英, 备份)
    记(f"开始：{英.name}，备份 {备份.name}，阈 {阈:.0%}，最多 {轮数} 轮")
    已挪, 已缩, 跳过 = {}, set(), set()
    for 轮 in range(1, 轮数 + 1):
        页 = 量(出版())
        坏 = [x for x in 页 if x[1] >= 阈 and x[2] and x[2] not in 跳过]
        总空 = sum(x[1] for x in 页 if x[1] > 0.05)
        记(f"第 {轮} 轮：{len(页)} 页，空白页（≥{阈:.0%}）{len(坏)} 处，空白合计 {总空:.2f} 页 —— " +
           "，".join(f"p{p} {空:.0%}→{名}" for p, 空, 名, *_ in 坏[:8]))
        if not 坏:
            break
        p, 空, 名, 剩, 图高, 注高 = 坏[0]
        # ① 先试缩图：缩到 ≥ 88% 就能放进本页剩余空间的，只缩不挪（不打乱图文顺序）
        if 名.startswith("Figure") and 图高 > 0 and 名 not in 已缩:
            比例 = (剩 - 注高 - 14) / 图高
            if 比例 >= 最小比例:
                r = 缩(名, min(比例, 1.0))
                if r:
                    已缩.add(名)
                    记(f"  p{p} 空 {空:.0%}：{名} 宽 {r[0]:g} → {r[1]:g} cm（缩到 {r[1] / r[0]:.0%}），留在本页")
                    continue
        # ② 缩不进去：往后挪一段，每张最多一次
        if 已挪.get(名, 0) >= 1:
            跳过.add(名); 记(f"  {名} 已挪过一次，不再挪（保节奏）"); continue
        原因 = 挪(名, 已挪)
        if 原因:
            跳过.add(名); 记("  " + 原因)
        else:
            记(f"  p{p} 空 {空:.0%}：{名} 往后挪一段（缩到 88% 也放不下）")
    记(f"结束。缩过：{sorted(已缩) or '无'}；挪过：{已挪 or '无'}；跳过：{sorted(跳过) or '无'}")


if __name__ == "__main__":
    main()
