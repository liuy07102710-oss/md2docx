# md2docx

**用 Markdown 去写论文怎么处理最后转换成word以及参考文献引用？**

## 要点：

1. 让导出的 Word 样式尽量贴近中文期刊论文模板。
2. **让 Markdown 里的引用关键词在 Word 中继续被 Zotero 插件识别和接管，方便后续阶段继续在 Word 里统一调整、刷新引用与更改参考文献格式。**

### 1. 提供pandoc转换时的参考模板与Lua过滤器

项目内置了一个参考模板 [templates/template_期刊论文.docx](templates/template_期刊论文.docx)，并通过 Lua 过滤器补充处理图题、表题、表格样式、字体颜色、行内代码等细节，使导出的 Word 更接近期刊写作习惯，而不是只得到一个“能打开”的通用 docx。

> 模板这部分主要参考这个开源项目[pandoc_docx_template](https://github.com/Achuan-2/pandoc_docx_template)，在此基础上只保留了md->docx，增加了两个Lua过滤器以及修改了一下word参考模板。

### 2. 支持 Zotero 字段化，而不只是 CSL 排版

常见的 `Pandoc + CSL` 参考文献管理，会把引用直接渲染成 Word 中的普通文本。这样虽然能得到结果，但转换成word后一般会再次调整局部格式或内容，此时如果突然要更改参考文献，就又要重返md中更改，再重新转换成word与调整局部内容，比较麻烦。

本项目解决思路：

- 先把 Markdown 转成普通 Word；
- 再把其中的引用关键词注入为 Zotero 可识别的 Word field code。

这样生成的 docx 可以继续被 Zotero 插件识别和接管，更适合“Markdown 写作，Word 终稿”的论文场景。

## 两种工作流的区别

| 方式 | 适合场景 | 引用在 Word 中的状态 |
| --- | --- | --- |
| `Pandoc + CSL` | 只想快速得到格式化参考文献 | 普通文本，便于查看，不便继续被 Zotero 接管 |
| `Pondoc + 字段化脚本` | 终稿还要在 Word 里继续调整引用 | Zotero field code，可继续被 Zotero 插件识别和刷新 |

用Pandoc+CSL的话还需要1个.bib参考文献加一个CSL样式文件，采用字段化注入脚本的话就不需要这两个文件，本地打开了Zotero并安装了Better BibTeX for Zotero插件即可。

## 目录结构

| 路径 | 说明 |
| --- | --- |
| [scripts/md2docx.py](scripts/md2docx.py) | 主转换脚本，负责 Markdown 转 docx |
| [scripts/inject_zotero_fieldcode_poc.ps1](scripts/inject_zotero_fieldcode_poc.ps1) | 第二步字段化脚本，把引用关键词注入为 Zotero 字段 |
| [markdown-to-docx.lua](markdown-to-docx.lua) | Lua 过滤器总入口 |
| [lua/](lua) | 各个细粒度 Pandoc Lua 过滤器 |
| [templates/template_期刊论文.docx](templates/template_期刊论文.docx) | 参考 Word 模板 |
| [tests/时间心理账户综述示例.md](tests/时间心理账户综述示例.md) | 一个可直接用于测试两种工作流的示例 Markdown |

## 环境依赖

### 基础依赖

- [Pandoc](https://pandoc.org/)
- Python 3.10+
- Python 包 `lxml`

安装 `lxml`：

```powershell
pip install lxml
```

### 使用 Zotero 字段化时的额外依赖

- 打开了 Zotero
- 安装了插件 Better BibTeX for Zotero
- Windows PowerShell

第二步脚本当前通过 Better BibTeX 暴露的本地 JSON-RPC 接口工作，默认地址为：

```text
http://127.0.0.1:23119/better-bibtex/json-rpc
```

## 快速开始

以下命令默认在 `tests` 文件夹中执行，也就是先进入：

```powershell
cd tests
```

### 1. 普通 Markdown 转 Word

直接使用项目自带模板：

```powershell
python ../scripts/md2docx.py 时间心理账户综述示例.md -o 时间心理账户_plain.docx
```

如果需要显式指定模板：

```powershell
python ../scripts/md2docx.py 时间心理账户综述示例.md -o 时间心理账户_plain.docx --reference ../templates/template_期刊论文.docx
```

### 2. 使用 CSL 直接导出参考文献

这条路线适合“直接得到格式化参考文献”的情况：

```powershell
pandoc 时间心理账户综述示例.md `
  -o 时间心理账户_CSL.docx `
  --from "markdown+raw_html" `
  --citeproc `
  --bibliography 时间心理账户.bib `
  --csl china-national-standard-gb-t-7714-2015-numeric.csl `
  --reference-doc ../templates/template_期刊论文.docx `
  --lua-filter ../markdown-to-docx.lua
```

### 3. 使用 Zotero 字段化路线

第一步，先生成普通 docx：

```powershell
python ../scripts/md2docx.py 时间心理账户综述示例.md -o 时间心理账户_plain.docx
```

第二步，把引用关键词注入为 Zotero 字段：

```powershell
powershell -ExecutionPolicy Bypass -File ../scripts/inject_zotero_fieldcode_poc.ps1 `
  -InputDocx 时间心理账户_plain.docx `
  -OutputDocx 时间心理账户_fieldcoded.docx
```

成功后，脚本会输出：

- 第一步是否完成；
- 第二步是否完成；
- 转换了多少个引用关键词；
- 去重后识别了多少条参考文献；
- 哪些 citekey 没有在 Better BibTeX 中找到。

## Markdown 中的引用写法

字段化脚本目前主要识别以下几种写法：citekey的内容来自于Zotero中Better BibTeX for Zotero插件里生成的（这个可以设置）。

- `[@citekey]`
- `[@A; @B]`
- `@citekey`
- `@A [@B; @C]`

示例可参考 [tests/时间心理账户综述示例.md](tests/时间心理账户综述示例.md)。

## 支持的格式细节

### Lua 过滤器

[markdown-to-docx.lua](markdown-to-docx.lua) 是总入口文件，按顺序加载多个 Lua 过滤器。

| 过滤器 | 作用 | 默认状态 |
| --- | --- | --- |
| `lua/paragraph-table-caption.lua` | 识别“表1 ...”单独段落，并把它转为紧邻表格的表题 | 开启 |
| `lua/paragraph-image-caption.lua` | 识别图片下一行的“图1 ...”段落，并把它写入图片标题 | 开启 |
| `lua/preserve_font_color.lua` | 尽量保留 Markdown / HTML 中的字体颜色 | 开启 |
| `lua/image-title-to-caption.lua` | 把图片 `title` 转成 Word 图注 | 开启 |
| `lua/add-inline-code.lua` | 让行内代码使用 `Inline Code` 样式 | 开启 |

### Markdown 与 Word 样式的对应关系

| Markdown / 结构 | Word 中对应样式 |
| --- | --- |
| 普通正文段落 | `Body Text` / `First Paragraph` |
| `#` 到 `######` | `Heading 1` 到 `Heading 6` |
| `表1 ...` + 紧邻表格 | `TableCaption` |
| Markdown 表格 / HTML 表格 | `Table` |
| 表格单元格文字 | `Compact` |
| 图片 | `Captioned Figure` |
| 图片 `title` 或“图1 ...”图注 | `Image Caption` |
| 引用块 `>` | `Block Text` |
| 代码块 | `Source Code` |
| 行内代码 ``code`` | `Inline Code` |

## 适合谁

这个项目特别适合：

- 用 Markdown 写中文社科论文的人；
- 想把 Pandoc 输出尽量贴近 Word 模板的人；
- 想在 Word 终稿阶段继续使用 Zotero 插件刷新引用的人。
