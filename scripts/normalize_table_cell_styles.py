#!/usr/bin/env python3
from __future__ import annotations

import argparse
import shutil
import tempfile
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}

DEFAULT_TARGET_STYLE = "Compact"
DEFAULT_SOURCE_STYLES = {"a0", "FirstParagraph"}


def iter_table_cell_paragraphs(root: etree._Element):
    for cell in root.xpath(".//w:tbl//w:tc", namespaces=NS):
        for paragraph in cell.xpath("./w:p", namespaces=NS):
            yield paragraph


def ensure_ppr(paragraph: etree._Element) -> etree._Element:
    ppr = paragraph.find(f"{{{W_NS}}}pPr")
    if ppr is None:
        ppr = etree.Element(f"{{{W_NS}}}pPr")
        paragraph.insert(0, ppr)
    return ppr


def set_paragraph_style(paragraph: etree._Element, target_style: str) -> bool:
    ppr = ensure_ppr(paragraph)
    pstyle = ppr.find(f"{{{W_NS}}}pStyle")
    if pstyle is None:
        pstyle = etree.Element(f"{{{W_NS}}}pStyle")
        ppr.insert(0, pstyle)

    current = pstyle.get(f"{{{W_NS}}}val")
    if current == target_style:
        return False

    if current is not None and current not in DEFAULT_SOURCE_STYLES:
        return False

    pstyle.set(f"{{{W_NS}}}val", target_style)
    return True


def normalize_table_cell_styles(docx_path: Path, *, target_style: str) -> int:
    changed = 0

    with tempfile.TemporaryDirectory() as tmp_dir_name:
        tmp_dir = Path(tmp_dir_name)
        unpack_dir = tmp_dir / "docx"
        with ZipFile(docx_path) as archive:
            archive.extractall(unpack_dir)

        document_xml = unpack_dir / "word" / "document.xml"
        root = etree.parse(str(document_xml))

        for paragraph in iter_table_cell_paragraphs(root.getroot()):
            if set_paragraph_style(paragraph, target_style):
                changed += 1

        if changed:
            root.write(
                str(document_xml),
                encoding="UTF-8",
                xml_declaration=True,
                standalone="yes",
            )

            rebuilt = tmp_dir / "rebuilt.docx"
            with ZipFile(rebuilt, "w", ZIP_DEFLATED) as archive:
                for path in unpack_dir.rglob("*"):
                    if path.is_file():
                        archive.write(path, path.relative_to(unpack_dir))

            shutil.copyfile(rebuilt, docx_path)

    return changed


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Normalize Word table cell paragraph styles to a target style such as Compact.",
    )
    parser.add_argument("docx_path", help="DOCX file to update in place")
    parser.add_argument(
        "--target-style",
        default=DEFAULT_TARGET_STYLE,
        help=f"Paragraph style to apply inside table cells. Default: {DEFAULT_TARGET_STYLE}",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    docx_path = Path(args.docx_path).expanduser()
    if not docx_path.is_file():
        raise SystemExit(f"DOCX file not found: {docx_path}")

    changed = normalize_table_cell_styles(
        docx_path,
        target_style=args.target_style,
    )
    print(f"Updated {changed} table-cell paragraphs in {docx_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
