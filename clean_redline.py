#!/usr/bin/env python3
"""
Post-process a Word-generated redline to remove obvious no-op tracked changes.

Microsoft Word Compare can mark unchanged text as a delete+insert pair when the
underlying DOCX structure differs, especially around hyperlinks, fields, and run
boundaries. This helper keeps substantive revisions but cleans common noise:
  - empty revision wrappers
  - whitespace-only insertions/deletions
  - adjacent insert/delete pairs with identical normalized text
"""

from __future__ import annotations

import copy
import datetime
import os
import re
import shutil
import sys
import tempfile
from pathlib import Path
from typing import Iterable
from zipfile import ZIP_DEFLATED, ZipFile
import xml.etree.ElementTree as ET


NAMESPACES = {
    "wpc": "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas",
    "cx": "http://schemas.microsoft.com/office/drawing/2014/chartex",
    "cx1": "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex",
    "cx2": "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex",
    "cx3": "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex",
    "cx4": "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex",
    "cx5": "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex",
    "cx6": "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex",
    "cx7": "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex",
    "cx8": "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "aink": "http://schemas.microsoft.com/office/drawing/2016/ink",
    "am3d": "http://schemas.microsoft.com/office/drawing/2017/model3d",
    "o": "urn:schemas-microsoft-com:office:office",
    "oel": "http://schemas.microsoft.com/office/2019/extlst",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "m": "http://schemas.openxmlformats.org/officeDocument/2006/math",
    "v": "urn:schemas-microsoft-com:vml",
    "wp14": "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "w10": "urn:schemas-microsoft-com:office:word",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "w16": "http://schemas.microsoft.com/office/word/2018/wordml",
    "w16du": "http://schemas.microsoft.com/office/word/2023/wordml/word16du",
    "w16sdtdh": "http://schemas.microsoft.com/office/word/2020/wordml/sdtdatahash",
    "w16sdtfl": "http://schemas.microsoft.com/office/word/2024/wordml/sdtformatlock",
    "w16se": "http://schemas.microsoft.com/office/word/2015/wordml/symex",
    "wpg": "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
    "wpi": "http://schemas.microsoft.com/office/word/2010/wordprocessingInk",
    "wne": "http://schemas.microsoft.com/office/word/2006/wordml",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
}

for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)

W = f"{{{NAMESPACES['w']}}}"
MC = f"{{{NAMESPACES['mc']}}}"
R = f"{{{NAMESPACES['r']}}}"
W14 = f"{{{NAMESPACES['w14']}}}"
INS = W + "ins"
DEL = W + "del"
TEXT = W + "t"
DEL_TEXT = W + "delText"
IGNORABLE = MC + "Ignorable"
URI_TO_PREFIX = {uri: prefix for prefix, uri in NAMESPACES.items()}
CHECKBOX_CHARS = {"☐", "☒"}

PROTECTED_TAGS = {
    W + "drawing",
    W + "object",
    W + "pict",
    W + "tbl",
    W + "sdt",
    W + "fldChar",
    W + "instrText",
    W + "bookmarkStart",
    W + "bookmarkEnd",
    W + "commentRangeStart",
    W + "commentRangeEnd",
    W + "commentReference",
    f"{{{NAMESPACES['v']}}}shape",
    f"{{{NAMESPACES['v']}}}imagedata",
    f"{{{NAMESPACES['o']}}}OLEObject",
}


def iter_text(element: ET.Element, tag: str) -> Iterable[str]:
    for child in element.iter(tag):
        if child.text:
            yield child.text


def revision_text(element: ET.Element) -> str:
    if element.tag == INS:
        return "".join(iter_text(element, TEXT))
    if element.tag == DEL:
        return "".join(iter_text(element, DEL_TEXT))
    return ""


def norm(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def loose_norm(text: str) -> str:
    return re.sub(r"\s+", "", text)


def same_text_or_spacing_only(left: str, right: str) -> bool:
    return norm(left) == norm(right) or loose_norm(left) == loose_norm(right)


def is_revision(element: ET.Element) -> bool:
    return element.tag in {INS, DEL}


def has_relationship_reference(element: ET.Element) -> bool:
    for node in element.iter():
        for attr_name, value in node.attrib.items():
            if value and attr_name.startswith(R):
                return True
    return False


def has_protected_content(element: ET.Element) -> bool:
    return any(node.tag in PROTECTED_TAGS for node in element.iter()) or has_relationship_reference(element)


def is_text_only_revision(element: ET.Element) -> bool:
    return is_revision(element) and not has_protected_content(element)


def is_unlinked_drawing_revision(element: ET.Element) -> bool:
    return (
        element.tag in {INS, DEL}
        and revision_text(element) == ""
        and any(node.tag == W + "drawing" for node in element.iter())
        and not has_relationship_reference(element)
    )


def has_visible_text_outside_revision(element: ET.Element) -> bool:
    def walk(node: ET.Element, in_revision: bool = False) -> bool:
        next_in_revision = in_revision or is_revision(node)
        if not next_in_revision and node.tag == TEXT and node.text:
            return True
        if not next_in_revision and node.tag == DEL_TEXT and node.text:
            return True
        return any(walk(child, next_in_revision) for child in list(node))

    return walk(element)


def nested_revision_payload(element: ET.Element) -> tuple[str, str, ET.Element] | None:
    if is_revision(element):
        if has_protected_content(element):
            return None
        return element.tag, revision_text(element), element

    if has_protected_content(element):
        return None

    if has_visible_text_outside_revision(element):
        return None

    revisions = []
    for node in element.iter():
        if is_revision(node):
            text = revision_text(node)
            if text:
                revisions.append((node.tag, text, node))

    if len(revisions) == 1:
        return revisions[0]
    return None


def accepted_children(element: ET.Element) -> list[ET.Element]:
    if element.tag == INS:
        return [copy.deepcopy(child) for child in list(element)]
    return []


def is_checkbox_sdt(element: ET.Element) -> bool:
    return element.tag == W + "sdt" and any(node.tag == W14 + "checkbox" for node in element.iter())


def first_text(element: ET.Element) -> str:
    for node in element.iter():
        if node.tag in {TEXT, DEL_TEXT} and node.text:
            return node.text
    return ""


def sdt_content_children(element: ET.Element) -> list[ET.Element]:
    for child in list(element):
        if child.tag == W + "sdtContent":
            return [copy.deepcopy(grandchild) for grandchild in list(child)]
    return []


def checkbox_sdt_replacement(parent: ET.Element, index: int) -> list[ET.Element] | None:
    child = parent[index]
    if not is_checkbox_sdt(child):
        return None

    text = first_text(child).strip()
    if text not in CHECKBOX_CHARS:
        return None

    if index + 1 < len(parent):
        sibling_text = first_text(parent[index + 1]).lstrip()
        if sibling_text[:1] in CHECKBOX_CHARS:
            return []

    return sdt_content_children(child)


def clean_parent(parent: ET.Element) -> int:
    changed = 0
    i = 0
    while i < len(parent):
        child = parent[i]

        if len(child):
            changed += clean_parent(child)

        replacement = checkbox_sdt_replacement(parent, i)
        if replacement is not None:
            parent[i : i + 1] = replacement
            changed += 1
            i += len(replacement)
            continue

        if not is_revision(child):
            i += 1
            continue

        text = revision_text(child)
        if text == "":
            if is_unlinked_drawing_revision(child):
                del parent[i]
                changed += 1
                continue
            if not is_text_only_revision(child):
                i += 1
                continue
            del parent[i]
            changed += 1
            continue

        if norm(text) == "":
            if not is_text_only_revision(child):
                i += 1
                continue
            replacement = accepted_children(child)
            parent[i : i + 1] = replacement
            changed += 1
            i += len(replacement)
            continue

        if i + 1 < len(parent):
            sibling = parent[i + 1]
            child_payload = nested_revision_payload(child)
            sibling_payload = nested_revision_payload(sibling)
            if child_payload and sibling_payload:
                child_kind, child_text, child_revision = child_payload
                sibling_kind, sibling_text, sibling_revision = sibling_payload
                source = child_revision if child_kind == INS else sibling_revision
                if (
                    child_kind != sibling_kind
                    and norm(child_text)
                    and same_text_or_spacing_only(child_text, sibling_text)
                    and source.tag == INS
                ):
                    replacement = accepted_children(source)
                    parent[i : i + 2] = replacement
                    changed += 1
                    i += len(replacement)
                    continue

        if i + 1 < len(parent) and is_revision(parent[i + 1]):
            sibling = parent[i + 1]
            sibling_text = revision_text(sibling)
            if (
                is_text_only_revision(child)
                and is_text_only_revision(sibling)
                and norm(text)
                and same_text_or_spacing_only(text, sibling_text)
                and child.tag != sibling.tag
            ):
                source = child if child.tag == INS else sibling
                replacement = accepted_children(source)
                parent[i : i + 2] = replacement
                changed += 1
                i += len(replacement)
                continue

        # INS before DEL where ins text == last run of del (del may have bookmarks).
        # Handles headings where Word Compare splits a no-op across a del with bookmarks.
        if (
            child.tag == INS
            and is_text_only_revision(child)
            and i + 1 < len(parent)
            and parent[i + 1].tag == DEL
            and not has_relationship_reference(parent[i + 1])
        ):
            sibling = parent[i + 1]
            result = _last_del_run_text(sibling)
            if result is not None:
                _, last_run, last_del_text = result
                if norm(last_del_text) and norm(last_del_text) == norm(text):
                    sibling.remove(last_run)
                    has_remaining = any(
                        r.tag == W + "r" and r.find(DEL_TEXT) is not None
                        for r in sibling
                    )
                    if not has_remaining:
                        del parent[i + 1]
                    replacement = accepted_children(child)
                    parent[i: i + 1] = replacement
                    changed += 1
                    i += len(replacement)
                    continue

        i += 1
    return changed


def text_revision_payloads(element: ET.Element) -> list[tuple[str, str, ET.Element]]:
    revisions = []
    for node in element.iter():
        if is_revision(node) and is_text_only_revision(node):
            text = revision_text(node)
            if norm(text):
                revisions.append((node.tag, text, node))
    return revisions


def has_visible_text_except_revision(element: ET.Element, ignored_revision: ET.Element) -> bool:
    def walk(node: ET.Element, in_revision: bool = False) -> bool:
        if node is ignored_revision:
            return False
        next_in_revision = in_revision or is_revision(node)
        if not next_in_revision and node.tag == TEXT and node.text:
            return True
        if not next_in_revision and node.tag == DEL_TEXT and node.text:
            return True
        return any(walk(child, next_in_revision) for child in list(node))

    return walk(element)


def replace_descendant(root: ET.Element, target: ET.Element, replacement: list[ET.Element]) -> bool:
    for index, child in enumerate(list(root)):
        if child is target:
            root[index : index + 1] = replacement
            return True
        if replace_descendant(child, target, replacement):
            return True
    return False


def clean_cross_paragraph_noops(root: ET.Element) -> int:
    changed = 0

    def walk(parent: ET.Element) -> None:
        nonlocal changed
        i = 0
        while i + 1 < len(parent):
            first = parent[i]
            second = parent[i + 1]
            if first.tag == W + "p" and second.tag == W + "p":
                first_payloads = text_revision_payloads(first)
                second_payloads = text_revision_payloads(second)
                if len(first_payloads) == 1 and len(second_payloads) == 1:
                    first_kind, first_text, first_revision = first_payloads[0]
                    second_kind, second_text, second_revision = second_payloads[0]
                    if (
                        first_kind == DEL
                        and second_kind == INS
                        and same_text_or_spacing_only(first_text, second_text)
                        and not has_visible_text_except_revision(second, second_revision)
                    ):
                        replace_descendant(first, first_revision, [])
                        replace_descendant(second, second_revision, accepted_children(second_revision))
                        changed += 1
            walk(first)
            i += 1
        if i < len(parent):
            walk(parent[i])

    walk(root)
    return changed


def paragraph_style(p: ET.Element) -> str | None:
    pPr = p.find(W + "pPr")
    if pPr is None:
        return None
    pStyle = pPr.find(W + "pStyle")
    if pStyle is None:
        return None
    return pStyle.get(W + "val")


def is_heading_paragraph(p: ET.Element) -> bool:
    style = paragraph_style(p)
    return style is not None and style.lower().startswith("heading")


def is_deletion_only_paragraph(p: ET.Element) -> bool:
    for child in p:
        tag = child.tag
        if tag in {W + "pPr", W + "bookmarkStart", W + "bookmarkEnd",
                   W + "commentRangeStart", W + "commentRangeEnd"}:
            continue
        if tag == DEL:
            continue
        for t in child.iter(TEXT):
            if t.text:
                return False
    return True


def only_accepted_ins(p: ET.Element) -> ET.Element | None:
    """Return the single <w:ins> if it is the only accepted (non-del) content in p."""
    ins_found = None
    for child in p:
        tag = child.tag
        if tag in {W + "pPr", W + "bookmarkStart", W + "bookmarkEnd",
                   W + "commentRangeStart", W + "commentRangeEnd",
                   W + "commentReference"}:
            continue
        if tag == DEL:
            continue
        if tag == INS:
            if ins_found is not None:
                return None
            if has_protected_content(child):
                return None
            ins_found = child
        else:
            # Normal run or other element — check for visible text
            for t in child.iter(TEXT):
                if t.text:
                    return None
    return ins_found


def clean_misplaced_heading_insertions(root: ET.Element) -> int:
    """Move a <w:ins> from a body paragraph into an immediately following empty heading.

    Word Compare sometimes deposits the heading text as an insertion in the last
    body paragraph of the previous section, leaving the heading paragraph itself
    empty. Detected pattern (within a window of 3 look-back paragraphs):

    - Para A (body): only visible accepted content is a single <w:ins> with text T
    - Para B (optional): deletion-only (no accepted content)
    - Para C (heading style): empty — no runs, no text
    """
    changed = 0

    def walk(parent: ET.Element) -> None:
        nonlocal changed
        paras = [child for child in parent if child.tag == W + "p"]
        para_indices = {id(p): i for i, p in enumerate(list(parent))}

        for idx, p in enumerate(paras):
            if not is_heading_paragraph(p):
                continue
            # Heading must be empty (no text anywhere, no ins/del)
            if any(t.text for t in p.iter(TEXT)):
                continue
            if p.find(".//" + INS) is not None or p.find(".//" + DEL) is not None:
                continue

            # Walk backward through up to 3 preceding paragraphs
            for back in range(1, 4):
                if idx - back < 0:
                    break
                candidate = paras[idx - back]
                if is_deletion_only_paragraph(candidate):
                    continue
                ins_elem = only_accepted_ins(candidate)
                if ins_elem is not None and revision_text(ins_elem):
                    # Move the <w:ins> from candidate into the heading paragraph
                    candidate.remove(ins_elem)
                    # Insert after <w:pPr> if present, otherwise at position 0
                    pPr = p.find(W + "pPr")
                    insert_pos = 0
                    for i, child in enumerate(p):
                        if child.tag == W + "pPr":
                            insert_pos = i + 1
                            break
                    p.insert(insert_pos, ins_elem)
                    changed += 1
                break

        for child in parent:
            if child.tag != W + "p":
                walk(child)

    walk(root)
    return changed


def _last_del_run_text(del_elem: ET.Element) -> tuple[ET.Element, ET.Element, str] | None:
    """Return (del_elem, last_run, text) for the last <w:r> with <w:delText> in del_elem."""
    last_run = None
    last_text = ""
    for r in del_elem:
        if r.tag != W + "r":
            continue
        dt = r.find(DEL_TEXT)
        if dt is not None and dt.text:
            last_run = r
            last_text = dt.text
    if last_run is None:
        return None
    return del_elem, last_run, last_text


def clean_paragraph_boundary_noop(root: ET.Element) -> int:
    """Clean a cross-paragraph boundary no-op where a word is deleted from the end
    of para N and inserted at the start of para N+1.

    Word Compare sometimes splits an unchanged word across a paragraph boundary:
    deleting it from the last run of a deletion at the end of para N and marking
    it as an insertion at the start of para N+1 (often a heading). When the texts
    match, both are removed and the word is treated as unchanged.
    """
    changed = 0

    def walk(parent: ET.Element) -> None:
        nonlocal changed
        i = 0
        while i + 1 < len(parent):
            first = parent[i]
            second = parent[i + 1]
            if first.tag == W + "p" and second.tag == W + "p":
                # Find the LAST top-level <w:del> in first
                last_del = None
                for child in first:
                    if child.tag == DEL:
                        last_del = child
                if last_del is not None:
                    result = _last_del_run_text(last_del)
                    if result is not None:
                        _, last_run, del_text = result
                        # Find the FIRST top-level <w:ins> in second
                        first_ins = None
                        for child in second:
                            if child.tag == INS:
                                first_ins = child
                                break
                        if first_ins is not None and not has_protected_content(first_ins):
                            ins_text = revision_text(first_ins)
                            if (
                                norm(del_text)
                                and norm(del_text) == norm(ins_text)
                            ):
                                # Remove last run from del; remove del if now empty
                                last_del.remove(last_run)
                                has_remaining = any(
                                    r.tag == W + "r" and r.find(DEL_TEXT) is not None
                                    for r in last_del
                                )
                                if not has_remaining:
                                    first.remove(last_del)
                                # Accept the insertion in second
                                ins_index = list(second).index(first_ins)
                                replacement = accepted_children(first_ins)
                                second[ins_index: ins_index + 1] = replacement
                                changed += 1
            walk(first)
            i += 1
        if i < len(parent):
            walk(parent[i])

    walk(root)
    return changed


def clean_misplaced_deletions_after_heading(root: ET.Element) -> int:
    """Move deletion-only paragraphs that appear between a clean heading and the
    section's first plain body paragraph to BEFORE the heading.

    Word Compare sometimes places old-content deletions after the new heading rather
    than before it, making deleted text appear to belong to the new section.

    Pattern:
    - Para H (heading): no ins, no del, has plain text
    - Para D1..DN: contiguous deletion-only paragraphs (with actual deletion content)
    - Para B: unchanged plain body text (the section's actual content)

    Fix: move D1..DN to immediately before H.
    """
    changed = 0

    def walk(parent: ET.Element) -> None:
        nonlocal changed
        made_change = True
        while made_change:
            made_change = False
            children = list(parent)
            for i, p in enumerate(children):
                if p.tag != W + "p":
                    continue
                if not is_heading_paragraph(p):
                    continue
                if p.find(".//" + INS) is not None or p.find(".//" + DEL) is not None:
                    continue
                if not norm("".join(t.text or "" for t in p.iter(TEXT))):
                    continue

                # Collect contiguous deletion-only paragraphs (must have del content)
                cluster = []
                j = i + 1
                while j < len(children):
                    cand = children[j]
                    if cand.tag != W + "p":
                        break
                    if cand.find(".//" + DEL) is not None and is_deletion_only_paragraph(cand):
                        cluster.append(cand)
                        j += 1
                    else:
                        break

                if not cluster:
                    continue

                # Validate: cluster must be followed by an unchanged plain body paragraph
                if j >= len(children) or children[j].tag != W + "p":
                    continue
                following = children[j]
                if (
                    following.find(".//" + INS) is not None
                    or following.find(".//" + DEL) is not None
                    or not norm("".join(t.text or "" for t in following.iter(TEXT)))
                ):
                    continue

                # Strip heading style from any cluster paragraph that has one —
                # the style was inherited from the wrong paragraph during Word Compare
                # alignment, and rendering it as a heading-styled deletion is misleading.
                for del_para in cluster:
                    if is_heading_paragraph(del_para):
                        pPr = del_para.find(W + "pPr")
                        if pPr is not None:
                            pStyle = pPr.find(W + "pStyle")
                            if pStyle is not None:
                                pPr.remove(pStyle)

                # Move cluster to before the heading
                for del_para in cluster:
                    parent.remove(del_para)
                new_children = list(parent)
                heading_pos = new_children.index(p)
                for offset, del_para in enumerate(cluster):
                    parent.insert(heading_pos + offset, del_para)
                changed += len(cluster)
                made_change = True
                break  # Restart scan after modification

        for child in parent:
            if child.tag != W + "p":
                walk(child)

    walk(root)
    return changed


def clean_heading_full_ins_noop(root: ET.Element) -> int:
    """Clean a no-op where a heading's only accepted content is <w:ins>T</w:ins> and a
    nearby following paragraph's last del run also has text T.

    Word Compare sometimes marks an unchanged heading as <w:ins> when paragraph counts
    differ, while placing the corresponding deletion a few paragraphs later (often
    another heading paragraph with old content). Detected pattern (forward window of 3):

    - Para A (heading): only visible accepted content is a single <w:ins>T</w:ins>
    - Para B (optional): deletion-only paragraphs — skip and keep scanning
    - Para C: has a top-level <w:del> whose LAST <w:r> has text T
    """
    changed = 0

    def walk(parent: ET.Element) -> None:
        nonlocal changed
        paras = [child for child in parent if child.tag == W + "p"]

        for idx, p in enumerate(paras):
            if not is_heading_paragraph(p):
                continue
            ins_elem = only_accepted_ins(p)
            if ins_elem is None:
                continue
            ins_text = revision_text(ins_elem)
            if not norm(ins_text):
                continue

            for fwd in range(1, 4):
                if idx + fwd >= len(paras):
                    break
                target = paras[idx + fwd]

                last_del = None
                for child in target:
                    if child.tag == DEL:
                        last_del = child

                if last_del is not None:
                    result = _last_del_run_text(last_del)
                    if result is not None:
                        _, last_run, del_text = result
                        if norm(del_text) == norm(ins_text):
                            ins_index = list(p).index(ins_elem)
                            replacement = accepted_children(ins_elem)
                            p[ins_index: ins_index + 1] = replacement
                            last_del.remove(last_run)
                            has_remaining = any(
                                r.tag == W + "r" and r.find(DEL_TEXT) is not None
                                for r in last_del
                            )
                            if not has_remaining:
                                target.remove(last_del)
                            changed += 1
                            break

                if not is_deletion_only_paragraph(target):
                    break

        for child in parent:
            if child.tag != W + "p":
                walk(child)

    walk(root)
    return changed


def clean_document_xml(xml_bytes: bytes) -> tuple[bytes, int]:
    root = ET.fromstring(xml_bytes)
    changed_ignorable = clean_ignorable_namespaces(root)
    total_changed = 0
    while True:
        changed = (
            clean_parent(root)
            + clean_cross_paragraph_noops(root)
            + clean_misplaced_heading_insertions(root)
            + clean_paragraph_boundary_noop(root)
            + clean_heading_full_ins_noop(root)
            + clean_misplaced_deletions_after_heading(root)
        )
        if changed == 0:
            break
        total_changed += changed
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), total_changed + changed_ignorable


def clean_ignorable_namespaces(root: ET.Element) -> int:
    ignorable = root.attrib.get(IGNORABLE)
    if not ignorable:
        return 0

    used_prefixes = set()
    for element in root.iter():
        names = [element.tag, *element.attrib.keys()]
        for name in names:
            if not name.startswith("{"):
                continue
            uri = name[1:].split("}", 1)[0]
            prefix = URI_TO_PREFIX.get(uri)
            if prefix:
                used_prefixes.add(prefix)

    old_prefixes = ignorable.split()
    new_prefixes = [prefix for prefix in old_prefixes if prefix in used_prefixes]
    new_ignorable = " ".join(new_prefixes)
    if new_ignorable == ignorable:
        return 0

    if new_ignorable:
        root.set(IGNORABLE, new_ignorable)
    else:
        del root.attrib[IGNORABLE]
    return 1


def rewrite_docx(path: Path) -> int:
    with ZipFile(path, "r") as zin:
        infos = zin.infolist()
        document_xml = zin.read("word/document.xml")
        cleaned_xml, changed = clean_document_xml(document_xml)

        if changed == 0:
            return 0

        fd, tmp_name = tempfile.mkstemp(suffix=".docx", dir=str(path.parent))
        os.close(fd)
        tmp_path = Path(tmp_name)
        try:
            with ZipFile(tmp_path, "w", ZIP_DEFLATED) as zout:
                for info in infos:
                    data = cleaned_xml if info.filename == "word/document.xml" else zin.read(info.filename)
                    zout.writestr(info, data)
            shutil.move(str(tmp_path), str(path))
        finally:
            if tmp_path.exists():
                tmp_path.unlink()

    return changed


def main() -> int:
    if len(sys.argv) != 2:
        print(f"Usage: {sys.argv[0]} <redline.docx>", file=sys.stderr)
        return 2

    path = Path(sys.argv[1])
    if not path.is_file():
        print(f"Error: file not found: {path}", file=sys.stderr)
        return 1

    changed = rewrite_docx(path)
    print(f"Cleaned {changed} redundant redline change wrapper(s).")

    log_path = Path("/tmp/redline_clean.log")
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with log_path.open("a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {path.name}: cleaned {changed} artifact(s)\n")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
