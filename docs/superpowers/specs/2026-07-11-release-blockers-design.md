# Release Blockers 2-4 Design

## Scope

Fix three confirmed release blockers without changing the public command-line workflows:

1. Resolve Markdown resources relative to the input Markdown file.
2. Inject Zotero fields into formatted or split Word runs without regex-based OOXML parsing.
3. Stop conversion when a required Lua filter module cannot be loaded.

Licensing, packaging, CI, and unrelated documentation improvements are out of scope.

## Resource Resolution

`scripts/md2docx.py` will pass the resolved parent directory of the input Markdown file to Pandoc as the default `--resource-path`. User-supplied Pandoc arguments after `--` remain later in the command and may override the default. Both the default Markdown-to-HTML path and `--direct` path will use the same resource path.

Acceptance criteria:

- Running the converter from the repository root against `tests/时间心理账户综述示例.md` embeds the image.
- Existing invocations from inside `tests/` continue to work.

## XML-Aware Zotero Injection

The PowerShell script will load `word/document.xml` as an XML DOM and operate on WordprocessingML nodes. Citation discovery will inspect text nodes rather than scan serialized XML. Existing Zotero field instructions will be excluded from discovery and replacement.

For each paragraph, the injector will concatenate eligible `w:t` text in document order and map character offsets back to their text and run nodes. Citation markers will be located in the combined paragraph text, processed from right to left, and replaced with field nodes. Text before and after a marker remains in its original runs. Empty runs created by removal will be deleted. The visible result run will copy the first matched run's `w:rPr`, so a bold or italic citation remains visibly formatted.

The existing citation parsing and Better BibTeX payload generation remain unchanged. After replacement, the script will rescan eligible text. Any remaining citation marker that was present in the input causes a non-zero exit and lists the unresolved markers. Missing Better BibTeX keys continue to be reported as warnings and remain visible, matching current behavior; they are not mistaken for structural replacement failures.

Acceptance criteria:

- Plain and bold citation markers are both converted.
- A marker split across multiple Word runs is converted.
- Prefix and suffix text in the same runs is preserved.
- Existing field instructions are not converted again.
- Structurally unhandled markers prevent a false success message.
- The output remains a valid DOCX ZIP and contains no converted citekey text.

## Required Lua Modules

Every module listed in `markdown-to-docx.lua` is required. A failed `require` will raise an error containing the module name and original Lua error. Pandoc must exit non-zero instead of generating a partially filtered document.

Acceptance criteria:

- The bundled filter loads normally.
- Running a copied entry filter without its `lua/` directory fails and names the missing module.

## Tests

Tests will use Python's standard `unittest` framework to avoid adding a test dependency. They will run real Pandoc and PowerShell processes, generate temporary DOCX files, and inspect OOXML with the existing `lxml` dependency. A local HTTP server will emulate the Better BibTeX JSON-RPC response for deterministic Zotero tests.

Test files will be isolated under `tests/` and use temporary directories. They will not overwrite committed example documents or require a running Zotero instance.

## Error Handling And Compatibility

- Python 3.10+ and Windows PowerShell remain supported.
- The Zotero command-line parameters and `MD2DOCX_BBT_ENDPOINT` override remain unchanged.
- Temporary files and mock servers are cleaned up even when tests fail.
- Production changes are limited to the three affected scripts plus focused tests.
