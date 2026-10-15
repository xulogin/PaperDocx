# Zotero 活引用：那三个文件怎么来

docx 里的引用要是**真的 Zotero 字段**（作者点 Refresh 能重排、能换样式），
而不是死文本，需要三样东西。

---

## 1. `references.json` —— CSL-JSON 元数据

Zotero 里选中这篇论文的文献集合 → 右键 → **导出条目** → 格式选 **CSL JSON**。

每条长这样，`id` 就是你在正文里用的标记：

```json
{
  "id": "ref1",
  "type": "article-journal",
  "title": "DART-lux: An unbiased and rapid Monte Carlo radiative transfer method…",
  "container-title": "Remote Sensing of Environment",
  "author": [{"family": "Wang", "given": "Yingjie"}, …],
  "issued": {"date-parts": [[2022, 6]]},
  "volume": "274",
  "page": "112973",
  "DOI": "10.1016/j.rse.2022.112973"
}
```

> ★ **DOI 字段留着。** 不出 DOI 是样式的事，不是删数据（见下面第 3 节）。

> ★ 导出的 json 里可能带 **HTML 实体**（`Photogrammetric Engineering &amp; Remote Sensing`）。
> 喂给 pandoc 之前要脱一层，否则参考文献里会印出 `&amp;` 四个字母。实测见过。

---

## 2. `zot_match_index.json` —— refN ↔ Zotero 条目

```json
{
  "ref1": {
    "itemKey": "6WRZWANT",
    "itemID": 2827,
    "uri": "http://zotero.org/users/local/PyJAb2jl/items/6WRZWANT"
  },
  …
}
```

**`uri` 是关键**：Zotero Refresh 时按 URI 直接命中条目，不会弹「请选择条目」。
`itemID` 填进字段的 `citationItems[].id`。

怎么拿：读 Zotero 的 `zotero.sqlite`（**复制一份到临时目录再读，别碰原库**），
`items` 表有 `itemID` 和 `key`，`libraryID` 对应的 URI 前缀在 `libraries` 表里。
按题名 / DOI 和 `references.json` 对上就行。

---

## 3. 引用样式 `.csl`

Zotero 自带的样式在 `<Zotero 数据目录>/styles/`。
投 Elsevier 就用 `elsevier-harvard.csl`。

### 不出 DOI：把 `access` 宏掏空

```xml
<!-- 原来 -->
<macro name="access">
  <choose>
    <if variable="DOI">
      <text variable="DOI" prefix="https://doi.org/"/>
    </if>
    …
  </choose>
</macro>

<!-- 改成 -->
<macro name="access"><text value=""/></macro>
```

**Zotero 库里的 DOI 一个都不用删。** 样式不去印它而已。

### 改完要装进 Zotero，而且要满足两个条件

```xml
<info>
  <title>Elsevier Harvard (no DOI)</title>
  <id>http://www.zotero.org/styles/elsevier-harvard-nodoi</id>
  <link href="http://www.zotero.org/styles/elsevier-harvard-nodoi" rel="self"/>
```

1. **`<id>` 和 `rel="self"` 的 href 必须一致**，不一致 Zotero 拒装
2. 另存到 `<Zotero 数据目录>/styles/elsevier-harvard-nodoi.csl`
3. ★ **重启 Zotero 才认得**

> ⛔ **不装的后果**：docx 里预渲染的文献表确实没 DOI，
> 但作者一在 Word 里点 Refresh，Zotero 会按它自己装着的 `elsevier-harvard` 重排，
> **DOI 全回来了**。交付时一定要告诉作者「重启一次 Zotero」。

---

## 字段长什么样

Zotero 的行内引用在 docx 里是 **5 个 run 的复杂字段**：

```xml
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText xml:space="preserve"> ADDIN ZOTERO_ITEM CSL_CITATION {…JSON…} RND xxxx </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
<w:r><w:t>(Wang et al., 2022)</w:t></w:r>      ← 显示文本，Refresh 时被重写
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

JSON 里每个 `citationItems` 要有 `id` / `uris` / `itemData` 三样。

### 解析时的两个坑

1. **`instrText` 会被 Word 拆到多个 run 里**（CSL JSON 动辄几 KB）。
   解析必须把一个字段内**所有** `instrText` 拼起来，只读第一个会拿到截断的 JSON。
2. **字段可以嵌套**。要靠 begin/end 计数配对，不能见到第一个 end 就收。

---

## 书目字段（参考文献表）

```xml
<!-- 塞进第一条的开头 -->
<w:r><w:fldChar w:fldCharType="begin"/></w:r>
<w:r><w:instrText> ADDIN ZOTERO_BIBL {"uncited":[],"omitted":[],"custom":[]} CSL_BIBLIOGRAPHY </w:instrText></w:r>
<w:r><w:fldChar w:fldCharType="separate"/></w:r>
   … 条目段落 …
<!-- 挂在最后一条的末尾 -->
<w:r><w:fldChar w:fldCharType="end"/></w:r>
```

> ★ 别另起两个空段落去放 begin 和 end —— 那样「References」标题和第一条之间会空一大截。

条目本身用 `pandoc --citeproc --csl=xxx.csl` 生成再搬过来。
搬完要**拆掉 `w:hyperlink`**（citeproc 会把题名做成蓝色超链接），只留里面的 run，
并清掉 `w:color` / `w:u` / `w:rStyle`。

---

## 让 Zotero 认领这份文档

`docProps/custom.xml` 里要有 `ZOTERO_PREF_1`：

```xml
<data data-version="3" zotero-version="10.0">
  <session id="随便一个稳定的字符串"/>
  <style id="http://www.zotero.org/styles/elsevier-harvard-nodoi"
         hasBibliography="1" bibliographyStyleHasBeenSet="1"/>
  <prefs>
    <pref name="fieldType" value="Field"/>
    <pref name="automaticJournalAbbreviations" value="false"/>
  </prefs>
</data>
```

还要在 `[Content_Types].xml` 加 Override、`_rels/.rels` 加 Relationship。

> ★ **python-docx 存盘后这个部件不一定还在**，掉了 Zotero 就不认这份文档了。
> 每次写完都无脑补一遍（幂等）。
