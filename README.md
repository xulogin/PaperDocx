# PaperDocx

把 Markdown 论文排成**投稿级 Word**：三线表、公式双列隐藏边框、Zotero 活引用（可去 DOI）、
图表自动定宽、改动标红；md 和 docx **双向联动**，作者在 Word 里调的样式不会被冲掉。

> pandoc 负责把 md 的结构变成 Word 原生对象（OMML 公式、表、图），
> python-docx 负责套样式和接 Zotero。两边各干各擅长的，别互相代劳。

## 装

```bash
claude plugin marketplace add xulogin/PaperDocx
claude plugin install paperdocx@paperdocx
```

依赖：`pandoc`、`python-docx`、`lxml`、`Pillow`（Python 3.10+）。
验收要 Microsoft Word（COM 导 PDF）。

## 它解决什么

一份 17 000 词 / 23 图 / 12 表 / 14 公式 / 62 条文献的遥感论文，要：

- 排成 Elsevier 投稿稿的样子（页边距、字号、三线表、公式编号）
- 参考文献接 Zotero，作者能 Refresh、能换样式、**不出 DOI 但库里 DOI 不删**
- 作者在 Word 上改字 → 改动回到 md；在 md 上改 → 改动到 Word
- **作者会在 Word 里调样式，调完不能被下次重出冲掉**
- 每版的改动在 Word 里标红给作者看

## 里面有什么

| 文件 | 干什么 |
|---|---|
| `排版配置.py` | ★ 排版参数的**唯一来源**。改版式只改这一个文件 |
| `出Word.py` | md → 排版好的 docx（支持 `--标红对照` 和 `--图dpi`） |
| `块模型.py` | 块解析 + 同步串（md ↔ Word 的公共中间表示） |
| `docx读写.py` | 在 docx 上**只读写文字，样式一个都不碰** |
| `三方同步.py` | md ↔ docx 双向联动，冲突就停手 |
| `对比版本.py` | 两版六项对比体检：信息守恒 / 词数 / 惯例 / 可读性 / 术语 / 结构 |
| `开新版.py` | V(N) → V(N+1)，旧版冻住 |
| `渲染看图.ps1` | 真 Word 导 PDF + 转 PNG，交付前肉眼验收 |

参数和踩的坑全在 `skills/paperdocx/SKILL.md`，
Zotero 那三个文件怎么来在 `references/zotero接法.md`，
压词和评估在 `references/压词与评估.md`。

## 几条值钱的实测

- **主题字覆盖显式字体**：pandoc 的 Heading 样式带 `asciiTheme`，只设 `w:ascii` 没用，
  渲染出来正文衬线、标题全无衬线。必须先删 `*Theme` 再设实名。
- **Word 导 PDF 慢是图的体量**：原图 59 MB 跑 30 分钟没出来，300 dpi 21 MB 只要 12 秒。
  **21 MB 和 33 MB 之间有断崖**，不是线性。只影响导 PDF，改字滚动不受影响。
- **表格列宽别按 0.5 em 估**：Times New Roman 的大写字母是 0.70 em，
  按 0.5 估会把表头 `LAI` 拆成 `LA↵I`。
- **放不下时先收格内边距再缩字号**：顺序反了，12 张表有 5 张掉到 7–8.5 pt；
  换过来 11 张留在 9 pt。
- **`New-Object -ComObject Word.Application` 会接管已经在跑的 Word**。
  一次卡住之后，后面每次自动化导出都在那个卡住的实例上干等。
- **Word COM 放在当前 PowerShell 进程里会莫名卡住**，丢进 `Start-Job` 13 秒就完。

## License

MIT
