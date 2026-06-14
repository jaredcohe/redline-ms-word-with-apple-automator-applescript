#!/usr/bin/env python3
"""Normalize DOCX structures before Word Compare.

Runs only on temporary working copies. It converts Word content controls
(`w:sdt`) to their visible `w:sdtContent` children so Word Compare does not
report no-op changes for dropdown/content-control wrappers.
"""

from __future__ import annotations

import os
import shutil
import sys
import tempfile
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from lxml import etree


W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}
W = f"{{{W_NS}}}"


def normalize_xml(xml_bytes: bytes) -> tuple[bytes, int]:
    parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False)
    root = etree.fromstring(xml_bytes, parser)
    changed = 0

    # Deepest first avoids moving a parent before its children are inspected.
    for sdt in reversed(root.xpath(".//w:sdt", namespaces=NS)):
        content = sdt.find("w:sdtContent", namespaces=NS)
        parent = sdt.getparent()
        if content is None or parent is None:
            continue
        idx = parent.index(sdt)
        children = list(content)
        parent.remove(sdt)
        for offset, child in enumerate(children):
            content.remove(child)
            parent.insert(idx + offset, child)
        changed += 1

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True), changed


def rewrite_docx(path: Path) -> int:
    with ZipFile(path, "r") as zin:
        infos = zin.infolist()
        normalized_xml, changed = normalize_xml(zin.read("word/document.xml"))

        if changed == 0:
            return 0

        fd, tmp_name = tempfile.mkstemp(suffix=".docx", dir=str(path.parent))
        os.close(fd)
        tmp_path = Path(tmp_name)
        try:
            with ZipFile(tmp_path, "w", ZIP_DEFLATED) as zout:
                for info in infos:
                    data = normalized_xml if info.filename == "word/document.xml" else zin.read(info.filename)
                    zout.writestr(info, data)
            shutil.move(str(tmp_path), str(path))
        finally:
            if tmp_path.exists():
                tmp_path.unlink()

    return changed


def main() -> int:
    if len(sys.argv) != 2:
        print(f"Usage: {sys.argv[0]} <docx>", file=sys.stderr)
        return 2

    path = Path(sys.argv[1])
    if not path.is_file():
        print(f"Error: file not found: {path}", file=sys.stderr)
        return 1

    changed = rewrite_docx(path)
    print(f"Normalized {changed} content control(s).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
