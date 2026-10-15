# -*- coding: utf-8 -*-
"""Word 复杂字段（complex field）的构造与解析 —— Zotero 引文字段那一套。

`build_doc.py`（从 Markdown 生成新 docx）和 `docx_edit.py`（在已有 docx 上
原地增删引用）都用这里的原语，**只有一份实现**，不会两边漂移。

Zotero 的引文字段在 docx 里长这样（5 个 run）：

    <w:r><w:fldChar w:fldCharType="begin"/></w:r>
    <w:r><w:instrText xml:space="preserve"> ADDIN ZOTERO_ITEM CSL_CITATION {...} RND xxxx </w:instrText></w:r>
    <w:r><w:fldChar w:fldCharType="separate"/></w:r>
    <w:r><w:t>(Author, 2020)</w:t></w:r>          ← 显示文本，Refresh 时被重写
    <w:r><w:fldChar w:fldCharType="end"/></w:r>

两个解析上的坑：

1. **`instrText` 会被 Word 拆到多个 run 里。** Zotero 的 CSL JSON 动辄几 KB，
   Word 存盘时会切开。解析必须把一个字段内的所有 `instrText` **拼起来**，
   只读第一个会拿到截断的 JSON。
2. **字段可以嵌套**（字段里套字段）。要靠 begin/end 计数配对，
   不能见到第一个 end 就收。
"""
from __future__ import annotations

import copy
import json
import random
import re
import string

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

ADDIN_ITEM = "ADDIN ZOTERO_ITEM CSL_CITATION"
ADDIN_BIBL = "ADDIN ZOTERO_BIBL"


def rand_id(n=10):
    return "".join(random.choices(string.ascii_letters + string.digits, k=n))


# ---------------------------------------------------------------------------
# 行内引用的显示文本（Harvard 风；Refresh 时 Zotero 会按样式重写，
# 这里只是让没 Refresh 之前肉眼看着也对）
# ---------------------------------------------------------------------------
def year_of(csl):
    try:
        return str(csl["issued"]["date-parts"][0][0])
    except (KeyError, IndexError, TypeError):
        return "n.d."


def author_string(csl):
    authors = csl.get("author") or csl.get("editor") or []
    if not authors:
        return csl.get("title", "Anon.")
    families = [a.get("family", a.get("literal", "?")) for a in authors]
    if len(families) == 1:
        return families[0]
    if len(families) == 2:
        return "{} and {}".format(families[0], families[1])
    return "{} et al.".format(families[0])


def format_inline(csl_list):
    """[(csl), ...] -> "(Author, 2020; Other et al., 2021, 2023)" """
    by_author, order = {}, []
    for csl in csl_list:
        auth = author_string(csl)
        if auth not in by_author:
            by_author[auth] = []
            order.append(auth)
        by_author[auth].append(year_of(csl))
    parts = []
    for auth in order:
        years = sorted(set(by_author[auth]))
        parts.append("{}, {}".format(auth, ", ".join(years)))
    return "(" + "; ".join(parts) + ")"


# ---------------------------------------------------------------------------
# run 的字符格式
# ---------------------------------------------------------------------------
def default_rpr():
    """build_doc 的排版：Times New Roman + 宋体 + 11pt。生成新文档时用。"""
    rpr = OxmlElement("w:rPr")
    rfonts = OxmlElement("w:rFonts")
    rfonts.set(qn("w:ascii"), "Times New Roman")
    rfonts.set(qn("w:hAnsi"), "Times New Roman")
    rfonts.set(qn("w:eastAsia"), "宋体")
    rpr.append(rfonts)
    sz = OxmlElement("w:sz")
    sz.set(qn("w:val"), "22")
    rpr.append(sz)
    return rpr


def rpr_like(run_el):
    """抄一个已有 run 的字符格式。

    ★ 往别人写好的文档里插引用时必须用这个，不能用 default_rpr()，
    否则插进去的引用字体字号和周围正文不一致，一眼能看出来是补的。
    """
    if run_el is None:
        return default_rpr()
    src = run_el.find(W + "rPr")
    if src is None:
        return OxmlElement("w:rPr")
    return copy.deepcopy(src)


# ---------------------------------------------------------------------------
# 构造
# ---------------------------------------------------------------------------
def citation_instr(citation_items, display_text):
    """拼出 ` ADDIN ZOTERO_ITEM CSL_CITATION {json} RND xxxx `。

    citation_items 每项要有 id / uris / itemData —— `uris` 是关键，
    Zotero Refresh 时按 URI 直接命中条目，不弹「请选择条目」。
    """
    payload = {
        "citationID": rand_id(),
        "properties": {
            "formattedCitation": display_text,
            "plainCitation": display_text,
            "noteIndex": 0,
        },
        "citationItems": citation_items,
        "schema": "https://github.com/citation-style-language/schema/"
                  "raw/master/csl-citation.json",
    }
    return " {} {} RND {} ".format(
        ADDIN_ITEM, json.dumps(payload, ensure_ascii=False, separators=(",", ":")),
        rand_id(10))


def bibliography_instr():
    payload = {"uncited": [], "omitted": [], "custom": []}
    return " {} {} CSL_BIBLIOGRAPHY ".format(
        ADDIN_BIBL, json.dumps(payload, ensure_ascii=False))


def make_field_runs(instr, display, rpr=None):
    """造出一个完整复杂字段的 5 个 w:r，按顺序返回。"""
    runs = []
    for kind, body in (("begin", None), ("instr", instr), ("separate", None),
                       ("text", display), ("end", None)):
        r = OxmlElement("w:r")
        r.append(copy.deepcopy(rpr) if rpr is not None else default_rpr())
        if kind == "instr":
            el = OxmlElement("w:instrText")
            el.set(qn("xml:space"), "preserve")
            el.text = body
            r.append(el)
        elif kind == "text":
            el = OxmlElement("w:t")
            el.set(qn("xml:space"), "preserve")
            el.text = body or ""
            r.append(el)
        else:
            el = OxmlElement("w:fldChar")
            el.set(qn("w:fldCharType"), kind)
            r.append(el)
        runs.append(r)
    return runs


def append_field(paragraph, instr, display, rpr=None):
    for r in make_field_runs(instr, display, rpr):
        paragraph._p.append(r)


def insert_field_after(run_el, instr, display, rpr=None):
    """把字段插在某个 run 之后。rpr 不给就抄那个 run 的格式。"""
    if rpr is None:
        rpr = rpr_like(run_el)
    anchor = run_el
    for r in make_field_runs(instr, display, rpr):
        anchor.addnext(r)
        anchor = r
    return anchor


# ---------------------------------------------------------------------------
# 解析
# ---------------------------------------------------------------------------
class Field(object):
    """文档里一个已存在的复杂字段。"""

    def __init__(self, paragraph, runs, instr, display):
        self.paragraph = paragraph
        self.runs = runs          # 组成这个字段的全部 w:r，按文档顺序
        self.instr = instr        # 拼好的完整 instrText
        self.display = display    # 显示文本

    @property
    def kind(self):
        s = self.instr.strip()
        if s.startswith(ADDIN_ITEM):
            return "citation"
        if s.startswith(ADDIN_BIBL):
            return "bibliography"
        return "other"

    def payload(self):
        """把 instrText 里的 JSON 抠出来解析。解析不了返回 None。"""
        m = re.search(r"\{.*\}", self.instr, re.S)
        if not m:
            return None
        try:
            return json.loads(m.group(0))
        except ValueError:
            return None

    def item_keys(self):
        """这个字段引了哪些 Zotero itemKey（从 uris 里取）。"""
        p = self.payload() or {}
        out = []
        for ci in p.get("citationItems") or []:
            for u in ci.get("uris") or []:
                k = u.rstrip("/").rsplit("/", 1)[-1]
                if k:
                    out.append(k)
        return out

    def titles(self):
        p = self.payload() or {}
        return [(ci.get("itemData") or {}).get("title") or ""
                for ci in (p.get("citationItems") or [])]

    def remove(self):
        for r in self.runs:
            parent = r.getparent()
            if parent is not None:
                parent.remove(r)

    def set_display(self, text):
        """改显示文本。Refresh 会重写它，但改了至少肉眼看着对。"""
        for r in self.runs:
            t = r.find(W + "t")
            if t is not None:
                t.text = text
                return True
        return False


def _fld_char_type(run_el):
    fc = run_el.find(W + "fldChar")
    if fc is None:
        return None
    return fc.get(qn("w:fldCharType"))


def iter_fields(paragraph):
    """扫一个段落，把里面的复杂字段逐个 yield 出来。

    靠 begin/end 计数配对（字段可以嵌套），并且把一个字段内
    **所有** instrText 拼起来 —— Word 会把长 JSON 拆开存。
    """
    runs = list(paragraph._p.findall(W + "r"))
    depth, cur, instr_parts, display_parts = 0, [], [], []
    for r in runs:
        ct = _fld_char_type(r)
        if ct == "begin":
            depth += 1
            if depth == 1:
                cur, instr_parts, display_parts = [], [], []
            cur.append(r)
            continue
        if depth == 0:
            continue
        cur.append(r)
        if ct == "end":
            depth -= 1
            if depth == 0:
                yield Field(paragraph, list(cur),
                            "".join(instr_parts), "".join(display_parts))
                cur, instr_parts, display_parts = [], [], []
            continue
        it = r.find(W + "instrText")
        if it is not None and it.text:
            instr_parts.append(it.text)
        t = r.find(W + "t")
        if t is not None and t.text:
            display_parts.append(t.text)


def all_paragraphs(doc):
    """正文段落 + 表格单元格里的段落 + 页眉页脚。引用可能藏在表格里。"""
    seen = set()

    def emit(p):
        if id(p._p) in seen:
            return
        seen.add(id(p._p))
        yield_list.append(p)

    yield_list = []
    for p in doc.paragraphs:
        emit(p)
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    emit(p)
                for t2 in cell.tables:
                    for r2 in t2.rows:
                        for c2 in r2.cells:
                            for p2 in c2.paragraphs:
                                emit(p2)
    for sec in doc.sections:
        for part in (sec.header, sec.footer,
                     getattr(sec, "first_page_header", None),
                     getattr(sec, "first_page_footer", None),
                     getattr(sec, "even_page_header", None),
                     getattr(sec, "even_page_footer", None)):
            if part is None:
                continue
            try:
                for p in part.paragraphs:
                    emit(p)
            except Exception:
                pass
    return yield_list


def find_zotero_fields(doc):
    """全文的 Zotero 字段，按文档顺序。返回 [(序号, Field), ...]"""
    out = []
    for p in all_paragraphs(doc):
        for f in iter_fields(p):
            if f.kind in ("citation", "bibliography"):
                out.append(f)
    return list(enumerate(out, 1))
