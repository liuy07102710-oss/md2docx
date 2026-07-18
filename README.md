# md2docx

把 Markdown 论文转换成更接近期刊格式的 Word，并在需要时把文献引用注入为 Word-Zotero插件可识别和刷新的 Word 字段。

```text
Markdown
  ↓
Pandoc 负责转换结构
  ↓
reference.docx 提供 Word 样式
  ↓
Lua 过滤器补充图题、表题等细节
  ↓
DOCX
  ↓
inject_zotero_fieldcode_poc.ps1 负责引用字段化被word-zotero插件接管
  ↓
DOCX
```

## 快速使用

下面的命令默认在项目根目录执行。使用时可以将要转换的文件放到tests文件夹里并更改对应命令即可。

[视频介绍](https://www.bilibili.com/video/BV1bUKF6wEPT/?spm_id_from=333.1387.homepage.video_card.click&vd_source=99f1f4cbd5c2d016da403ccccb1f09f6)

### 方式一. 只做 Markdown 转 Word

```powershell
python scripts/md2docx.py "tests/时间心理账户综述示例.md" -o "tests/时间心理账户_plain.docx"
```

这条命令适合：先把 Markdown 转成样式更可控的 Word，后续还要在 Word 中继续调格式，还没有处理文献引用。

### 方式二. 直接用 CSL 生成参考文献

```powershell
pandoc "tests/时间心理账户综述示例.md" `
  -o "tests/时间心理账户_CSL.docx" `
  --from "markdown+raw_html" `
  --citeproc `
  --bibliography "tests/时间心理账户.bib" `
  --csl "tests/china-national-standard-gb-t-7714-2015-numeric.csl" `
  --reference-doc "templates/template_期刊论文.docx" `
  --lua-filter "markdown-to-docx.lua"
```

这条路线适合：直接得到带参考文献的 Word，生成后基本不再改引用，不需要在 Word 里继续让 Zotero 接管这些引用。

### （推荐）方式三. 引用注入为 Zotero 字段

第一步，先生成普通 Word：

```powershell
python scripts/md2docx.py "tests/时间心理账户综述示例.md" -o "tests/时间心理账户_plain.docx"
```

第二步，把 citekey 注入为 Zotero 字段：

```powershell
powershell -ExecutionPolicy Bypass -File "scripts/inject_zotero_fieldcode_poc.ps1" `
  -InputDocx "tests/时间心理账户_plain.docx" `
  -OutputDocx "tests/时间心理账户_fieldcoded.docx"
```

这条路线适合：Markdown 完成主要写作，Word 用来做终稿，进入 Word 后使用zotero插件刷新引用、切换样式、插入参考文献列表。

## 这个项目解决什么问题

现在无论是网页 AI、桌面 Agent，还是 Obsidian、Typora 和 VS Code，Markdown 都是一种非常灵活的写作格式。但在中文学术写作，特别是社会科学期刊投稿中，最终往往还是要提交 Word。

对于markdown-->word，大部分的做法是采用pandoc进行转换，但或多或少转换的不完美，主要存在两个方面问题，一是word模板的选择，二是文献引用的方式。

针对第一方面，参考开源项目[pandoc_docx_template](https://github.com/Achuan-2/pandoc_docx_template)，在此基础上改动了一下，制作了[word模板](templates/template_期刊论文.docx)，并增加了几个lua过滤器。最后的效果在标题、正文、图片、表格、图题、表题、公式、代码块都贴近C刊的风格，且模板可根据具体期刊模板进行修改对应的样式（见下方语法样式对照表）。

针对第二方面，一般做法是采用CSL参考文献样式文件+bib参考文献，并在md开头写上yaml标签，这种做法的缺点是后续要更改参考文献的话比较麻烦。本项目通过字段化注入，可以让转换后的 Word 文件里面的参考文献引用被 word-Zotero 插件识别并管理。

## 环境依赖

### 基础转换依赖

- [Pandoc](https://pandoc.cn/installing.html)
- Python 3.10+
- Python 包 `lxml`

### Zotero 字段化的额外依赖

- Windows PowerShell
- 本机已打开 Zotero且已安装 Better BibTeX for Zotero

Better BibTeX 默认通过本地 JSON-RPC 接口提供文献查询：

```text
http://127.0.0.1:23119/better-bibtex/json-rpc
```

## Markdown 中的引用写法

字段化脚本当前主要识别以下几种写法：

- `[@citekey]`
- `[@A; @B]`
- `@citekey`
- `@A [@B; @C]`

示例可参考 [tests/时间心理账户综述示例.md](tests/时间心理账户综述示例.md)。


## 目录结构

| 路径                                                         | 说明                                      |
| ------------------------------------------------------------ | ----------------------------------------- |
| [scripts/md2docx.py](scripts/md2docx.py)                     | 主转换脚本，负责 Markdown 转 Word         |
| [scripts/inject_zotero_fieldcode_poc.ps1](scripts/inject_zotero_fieldcode_poc.ps1) | 字段化脚本，把 citekey 注入为 Zotero 字段 |
| [markdown-to-docx.lua](markdown-to-docx.lua)                 | Lua 过滤器总入口                          |
| [lua/](lua)                                                  | 各个细粒度 Pandoc Lua 过滤器              |
| [templates/template_期刊论文.docx](templates/template_期刊论文.docx) | 参考 Word 模板                            |
| [tests/时间心理账户综述示例.md](tests/时间心理账户综述示例.md) | 演示用 Markdown 示例                      |
| [tests/时间心理账户.bib](tests/时间心理账户.bib)             | CSL 路线使用的参考文献库                  |
| [tests/china-national-standard-gb-t-7714-2015-numeric.csl](tests/china-national-standard-gb-t-7714-2015-numeric.csl) | CSL 样式文件                              |

## Lua 过滤器

[markdown-to-docx.lua](markdown-to-docx.lua) 是总入口文件，按顺序加载多个 Lua 过滤器。

| 过滤器                            | 作用                                              | 默认状态 |
| --------------------------------- | ------------------------------------------------- | -------- |
| `lua/paragraph-table-caption.lua` | 识别“表1 ...”单独段落，并把它转为紧邻表格的表题   | 开启     |
| `lua/paragraph-image-caption.lua` | 识别图片下一行的“图1 ...”段落，并把它写入图片标题 | 开启     |
| `lua/preserve_font_color.lua`     | 尽量保留 Markdown / HTML 中的字体颜色             | 开启     |
| `lua/image-title-to-caption.lua`  | 把图片 `title` 转成 Word 图注                     | 开启     |
| `lua/add-inline-code.lua`         | 让行内代码使用 `Inline Code` 样式                 | 开启     |

## Markdown语法 与 Word 样式的对应关系

| Markdown                             | Word 中对应样式                 |
| ------------------------------------ | ------------------------------- |
| 普通正文段落                         | `Body Text` / `First Paragraph` |
| `#` 到 `######`                      | `Heading 1` 到 `Heading 6`      |
| `表1 ...` + 紧邻表格(表题)           | `TableCaption`                  |
| Markdown 表格 / HTML 表格            | `Table`                         |
| 表格单元格文字                       | `Compact`                       |
| 图片                                 | `Captioned Figure`              |
| 图片 `title` 或“图1 ...”图注（图题） | `Image Caption`                 |
| 引用块 `>`                           | `Block Text`                    |
| 代码块                               | `Source Code`                   |
| 行内代码 ``code``                    | `Inline Code`                   |


## 致谢

模板部分主要参考开源项目 [pandoc_docx_template](https://github.com/Achuan-2/pandoc_docx_template)，本项目在此基础上保留了 Markdown 到 Word 的工作流，并补充了图表题、表格样式和 Zotero 字段化处理。
