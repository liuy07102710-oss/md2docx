# md2docx

## 1. 过滤器说明

`markdown-to-docx.lua` 是总过滤器入口文件，用来按顺序加载本项目实际使用的多个 Lua 过滤器；Pandoc 转 Word 时通常直接挂这个文件，而不是逐个手写所有过滤器路径。

| 过滤器 | 作用 | 默认状态 |
| --- | --- | --- |
| `lua/paragraph-table-caption.lua` | 识别“表1 ...”这类单独段落，并把它转成紧邻表格的表题。 | 开启 |
| `lua/paragraph-image-caption.lua` | 识别图片下一行的“图1 ...”这类段落，并把它写入图片标题。 | 开启 |
| `lua/preserve_font_color.lua` | 尽量保留 Markdown/HTML 中的字体颜色。 | 开启 |
| `lua/image-title-to-caption.lua` | 把图片 `title` 转成 Word 图注。 | 开启 |
| `lua/add-inline-code.lua` | 让行内代码使用单独的 `Inline Code` 样式。 | 开启 |
| `lua/markdown-html-recognition.lua` | 处理部分 HTML 标签，如 `<sub>`、`<sup>`、`<img>`。 | 关闭 |
| `lua/image-title-to-caption-add-number.lua` | 给图片图注自动加编号。 | 关闭 |

## 2. Python 脚本处理逻辑

项目默认使用 [scripts/md2docx.py](/D:/Vibe_Project/md2docx/scripts/md2docx.py:1) 进行转换。

处理流程很简单：

1. 先把 Markdown 转成 HTML。
2. 再把 HTML 转成 Word `.docx`。
3. 转换时套用 Word 参考模板。
4. 同时加载 Lua 过滤器处理表题、图题、颜色、行内代码等细节。
5. 最后再做一次 docx 后处理，把表格单元格里的文字样式统一改成 `Compact`。

先转 HTML、再转 docx 的目的，是让 Pandoc 更容易保留 Markdown 里夹带的 HTML 内容。`lua/markdown-html-recognition.lua` 这个过滤器是补充处理局部 HTML 标签细节的，不和“先转 HTML”这一步重复。

如果命令里加了 `--direct`，则会跳过“先转 HTML”这一步，直接从 Markdown 转 Word。

## 3. Markdown 语法与 Word 样式对应关系

| Markdown / 结构 | Word 中对应样式 |
| --- | --- |
| 普通正文段落 | `Body Text` / `First Paragraph` |
| `#` 一级标题 | `Heading 1` |
| `##` 二级标题 | `Heading 2` |
| `###` 三级标题 | `Heading 3` |
| `####` 四级标题 | `Heading 4` |
| `#####` 五级标题 | `Heading 5` |
| `######` 六级标题 | `Heading 6` |
| `表1 ...` + 紧邻表格 | `TableCaption` |
| Markdown 表格 / HTML 表格 | `Table` |
| 表格单元格文字 | `Compact` |
| 图片 | `Captioned Figure` |
| 图片 `title` 或“图1 ...”图注 | `Image Caption` |
| 引用块 `>` | `Block Text` |
| 代码块 | `Source Code` |
| 行内代码 ``code`` | `Inline Code` |

## 4. Python 转换命令

### 4.1 推荐命令

```powershell
python "D:\Vibe_Project\md2docx\scripts\md2docx.py" "input.md" -o "output.docx" --reference "D:\Vibe_Project\md2docx\templates\template_期刊论文.docx"
```

### 4.2 补充说明

- 如果 Markdown 中混有较多原始 HTML，可加 `--from "markdown+raw_html"`，明确要求 Pandoc 处理原始 HTML。
- 如果想跳过“先转 HTML、再转 Word”这一步，可加 `--direct`，直接从 Markdown 转 Word。
- 如果不写 `--reference`，脚本会使用默认模板。

## 5. Zotero 字段化补充

### 5.1 Markdown 中引用关键词怎么写

当前脚本主要识别这些写法：

- `[@citekey]`
- `[@A; @B]`
- `@citekey`
- `@A [@B; @C]`

示例可参考 [zotero-fieldcode/sample-zotero-full.md](/D:/Vibe_Project/md2docx/zotero-fieldcode/sample-zotero-full.md:1)。

### 5.2 使用字段化脚本的命令

先把 Markdown 转成普通 Word：

```powershell
python "D:\Vibe_Project\md2docx\scripts\md2docx.py" "D:\Vibe_Project\md2docx\zotero-fieldcode\sample-zotero-full.md" -o "D:\Vibe_Project\md2docx\zotero-fieldcode\sample-zotero-full-plain.docx" --reference "D:\Vibe_Project\md2docx\templates\template_期刊论文.docx"
```

再把普通 Word 字段化：

```powershell
powershell -ExecutionPolicy Bypass -File "D:\Vibe_Project\md2docx\scripts\inject_zotero_fieldcode_poc.ps1" -InputDocx "D:\Vibe_Project\md2docx\zotero-fieldcode\sample-zotero-full-plain.docx" -OutputDocx "D:\Vibe_Project\md2docx\zotero-fieldcode\sample-zotero-full-fieldcoded.docx"
```

### 5.3 字段化脚本原理

这个脚本不是直接把 Markdown 变成 Zotero Word 字段，而是：

1. 先生成一个普通 `.docx`。
2. 再解包这个 `docx`，读取 `word/document.xml`。
3. 在 XML 里识别还保留下来的 `@citekey` / `[@citekey]` 标记。
4. 通过本地运行的 Zotero + Better BibTeX 查询 citekey 对应条目。
5. 把这些位置替换成 Zotero Word field code。
6. 最后重新打包成新的 `.docx`。

这样生成的 Word 文档可以被 Zotero 插件识别和接管，后续可在 Word 中刷新引用样式和参考文献。
