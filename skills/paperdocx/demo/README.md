# Demo：一篇把所有排版元素都用上的样稿

**内容完全虚构**（一篇关于低成本声学雨量计的假论文），只为演示排版。
数据、作者、机构、参考文献全是编的，和任何真实研究无关。

## 跑

```powershell
cd <这个 demo 目录>
$env:PYTHONIOENCODING='utf-8'
py -3.12 -X utf8 做demo素材.py            # 生成图和引用数据（固定种子，可复现）
py -3.12 -X utf8 ../scripts/出Word.py            # → C_英文稿V01_排版版.docx
powershell -ExecutionPolicy Bypass -File ..\scripts\渲染看图.ps1   # 真 Word 导 PDF + 转 PNG
```

6 页，Word 导 PDF 约 7 秒。

## 它演示了什么（也就是模板里有什么）

| 元素 | 在 demo 的哪 | 排成什么样 |
|---|---|---|
| 题名 | 第 1 页 | Times New Roman 16 pt 粗体居中 |
| 作者 + 上标编号 | 第 1 页 | 11 pt 居中，`^1,\*^` 渲染成真上标 |
| 单位 | 第 1 页 | 8.5 pt，悬挂缩进 0.55 cm，一行一个 |
| 摘要 / 关键词 | 第 1 页 | 正文 11 pt / 行距 1.5 / 两端对齐 |
| 三级标题 | 全文 | H2 13.5、H3 11.5、H4 11，都加粗，段前段后递减 |
| **行内公式** | 2.2 节 | `$A$`、`$\Delta t$` → OMML，跟着正文字号走 |
| **带编号公式** | 式 (1)(2) | 1×2 隐藏边框表：左公式居中、右编号居中 |
| **三线表** | 表 1、2、A1 | 上线粗 / 表头下线细 / 末行粗，其余无线；**每格水平+垂直居中**；表头加粗；9 pt |
| 表题 / 表注 | 表上下 | 表题 `**Table 1.**` 标签加粗；表注 9 pt |
| **图 + 图注** | 图 1、2 | 等比缩放到 `{width=14cm}`，图注 9.5 pt，图注跟着图不分家 |
| **Zotero 活引用** | 全文 7 处 | 真字段（`ADDIN ZOTERO_ITEM`），不是死文本 |
| **参考文献** | 第 5 页 | Zotero 书目字段，著者—年份，悬挂缩进，**黑字无超链接，不出 DOI** |
| 附录表 | 第 6 页 | 和正文表同一套三线表样式 |
| 页码 | 页脚 | `PAGE` 域，居中 |

> ★ **不出 DOI 不是删数据**：`引用数据/references.json` 里每条都带 `DOI` 字段，
> 是引用样式（`demo-author-date.csl` 的 `access` 宏被掏空）不去印它。

## 拿它当模板怎么用

1. 把 `C_英文稿V01_英文版.md` 复制成你自己的稿子，**保留结构**：
   题名 → `**Authors:**` → `**Affiliations:**` → `**Correspondence:**` → `## Abstract` → …
2. 图放 `figures/`，在 md 里 `![](figures/xxx.png){width=14cm}`，下一块写 `**Figure N.** 图注`
3. 公式写 `$$ ... \qquad\qquad (N) $$`，下一行可选 `<!--eq:标签-->`
4. 表用 markdown 管道表，上一块写 `**Table N.** 表题`，下一块可选 `Note: 表注`
5. 引用在正文里直接写渲染形式 `(Hollis et al., 2019)`，
   由 `引用映射.tsv` 映回 `refN`，出 Word 时换成 Zotero 字段
6. 换刊物：`py -3.12 -X utf8 ../scripts/探参考docx.py 那本刊物的稿子.docx --配置`，
   把吐出来的参数抄进 `../scripts/排版配置.py`

## 目录

```
demo/
  C_英文稿V01_英文版.md        样稿（所有元素都用上了）
  做demo素材.py               生成图和引用数据
  引用映射.tsv                refN ↔ 渲染引用
  figures/                   生成的两张图
  引用数据/                   references.json / zot_match_index.json / demo-author-date.csl
（脚本不在这儿，直接用上一层的 ../scripts/）
```
