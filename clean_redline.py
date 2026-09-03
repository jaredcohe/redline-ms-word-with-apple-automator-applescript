#!/usr/bin/env python3
"""
Post-process a Word-generated redline to remove obvious no-op tracked changes.

Microsoft Word Compare can mark unchanged text as a delete+insert pair when the
underlying DOCX structure differs, especially around hyperlinks, fields, and run
boundaries. This helper keeps substantive revisions but cleans common noise:
  - empty revision wrappers
  - whitespace-only insertions/deletions
  - adjacent insert/delete pairs with identical normalized text

It also normalizes body formatting on every run: body paragraphs (non-heading/title,
tables included) are forced to 10pt font and 1.15 line spacing via silent direct
formatting (see normalize_body_formatting).
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
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
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

# Hyperlink machinery. These appear in PROTECTED_TAGS because they must never be *dropped*,
# but a revision whose only "protected" content is a hyperlink is still safe to resolve by
# ACCEPTING it (accepted_children deep-copies the field verbatim). See is_textlike_revision.
HYPERLINK_TAGS = {
    W + "fldChar",
    W + "instrText",
    W + "fldSimple",
    W + "hyperlink",
}

# Bookmarks are protected because they must not be dropped, but a paragraph that only
# anchors bookmarks can still be removed once its marks are relocated to a neighbour.
BOOKMARK_TAGS = {
    W + "bookmarkStart",
    W + "bookmarkEnd",
}

# Property containers where a bare <w:ins>/<w:del> is a revision *marker*, not a wrapper
# around content: a tracked paragraph mark (pPr/rPr), or an inserted/deleted table row or
# cell. These are legitimately empty and must never be pruned as empty revision wrappers,
# or a real "this paragraph break was deleted" edit vanishes un-rejectably.
REVISION_MARKER_PARENTS = {
    W + "rPr",
    W + "trPr",
    W + "tcPr",
    W + "tblPr",
    W + "tblPrEx",
}

# Elements that represent an embedded picture (the logo). A tracked change whose only
# content is one of these is resolved so the image never renders as a tracked change.
IMAGE_TAGS = {
    W + "drawing",
    W + "pict",
    f"{{{NAMESPACES['v']}}}shape",
    f"{{{NAMESPACES['v']}}}imagedata",
}

# Embedded OLE objects (spreadsheets, equations, …) are NOT logos — a revision containing
# one is left alone even though it may render as a picture.
OLE_TAGS = {
    W + "object",
    f"{{{NAMESPACES['o']}}}OLEObject",
}

# OOXML child-element order (local names) for CT_PPr and CT_RPr. New children must be
# inserted in this order or Word flags the file for repair.
PPR_ORDER = (
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
    "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
    "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN",
    "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing",
    "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment",
    "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange",
)
RPR_ORDER = (
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
    "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish",
    "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight",
    "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang",
    "eastAsianLayout", "specVanish", "oMath",
)

# Body formatting targets (user house style): 10pt = sz/szCs val="20" (half-points);
# 1.15 line spacing = spacing line="276" lineRule="auto" (240 * 1.15).
BODY_FONT_HALF_PT = "20"
BODY_LINE = "276"
NON_BODY_STYLES = {"Title", "Subtitle"}

# Space kept above the document title when the leading spacer paragraphs are removed
# (16pt = the original template's Title/@w:after). See clean_leading_empty_paragraphs.
TITLE_GAP_AFTER = "320"

# How far apart (in flattened document order) a relocated deletion and its matching
# insertion may sit. See clean_relocated_table_noops.
RELOCATION_WINDOW = 4


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


def is_textlike_revision(element: ET.Element) -> bool:
    """Like is_text_only_revision, but tolerates hyperlinks (fields or <w:hyperlink>).

    Word Compare re-emits a mailto/URL hyperlink as an insert+delete pair whenever the two
    documents build the link differently (a HYPERLINK field vs. a <w:hyperlink> element), so
    a signature block's e-mail address shows as edited when only its whitespace changed.
    Such a revision is safe to resolve *by accepting the insertion*, because
    accepted_children deep-copies the whole field/hyperlink verbatim — nothing is lost.

    Everything else in PROTECTED_TAGS still blocks: drawings, OLE objects, tables, content
    controls, bookmarks and comment anchors. Relationship references are allowed only when
    they hang off a <w:hyperlink> (an image's r:embed still blocks).
    """
    if not is_revision(element):
        return False

    blocking = PROTECTED_TAGS - HYPERLINK_TAGS
    for node in element.iter():
        if node.tag in blocking:
            return False

    hyperlink_ids = {id(node) for node in element.iter(W + "hyperlink")}
    for node in element.iter():
        if id(node) in hyperlink_ids:
            continue
        for attr_name, value in node.attrib.items():
            if value and attr_name.startswith(R):
                return False
    return True


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


def del_to_unchanged_runs(del_elem: ET.Element) -> list[ET.Element]:
    """Convert a <w:del> into plain unchanged runs (delText -> t), preserving rPr.

    Used to turn a spurious deletion of unchanged text back into normal body text.
    """
    runs = []
    for child in list(del_elem):
        if child.tag != W + "r":
            continue
        run = copy.deepcopy(child)
        for dt in run.findall(DEL_TEXT):
            dt.tag = TEXT
        runs.append(run)
    return runs


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

        # A revision inside a property container marks the paragraph mark / row / cell
        # itself. It carries no text by design, so none of the wrapper rules below apply.
        if parent.tag in REVISION_MARKER_PARENTS:
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
                is_textlike_revision(child)
                and is_textlike_revision(sibling)
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


def paragraph_mark_rPr(p: ET.Element) -> ET.Element | None:
    """The paragraph mark's own run properties (w:pPr/w:rPr), if present."""
    pPr = p.find(W + "pPr")
    if pPr is None:
        return None
    return pPr.find(W + "rPr")


def paragraph_mark_revision(p: ET.Element) -> str | None:
    """INS or DEL if the paragraph mark itself is a tracked change, else None."""
    rPr = paragraph_mark_rPr(p)
    if rPr is None:
        return None
    if rPr.find(INS) is not None:
        return INS
    if rPr.find(DEL) is not None:
        return DEL
    return None


def paragraph_mark_has_revision(p: ET.Element) -> bool:
    """True if the paragraph mark itself is a tracked change (w:pPr/w:rPr/w:ins|w:del)."""
    return paragraph_mark_revision(p) is not None


def clean_empty_list_paragraphs(root: ET.Element) -> int:
    """Remove orphan empty list paragraphs that render a bare list marker.

    When paragraph counts differ between the two compared documents, Word Compare emits
    empty list paragraphs as alignment padding. They carry a <w:numPr> (so they render a
    bullet/number/letter marker, e.g. numId=60 -> "(a) (b) (c) ...") but hold no visible
    text and no deleted text, and their paragraph mark is not itself a tracked change, so
    they survive both Accept and Reject. The only way to remove them is to delete the
    <w:p> outright.

    A paragraph is removed only when ALL hold:
    - it has list numbering (pPr/numPr present),
    - it has no visible text and no deleted text (whitespace-only counts as empty),
    - it has no protected content (bookmarks, drawings, tables, fields, comment anchors,
      hyperlinks) — this preserves list paragraphs that anchor cross-reference bookmarks,
    - its paragraph mark is not itself a tracked change.

    Paragraphs whose content is a tracked deletion (e.g. "Intentionally omitted.") are
    kept, so real deletions stay visible in the redline.
    """
    changed = 0

    def is_empty_list_paragraph(p: ET.Element) -> bool:
        pPr = p.find(W + "pPr")
        if pPr is None or pPr.find(W + "numPr") is None:
            return False
        if norm("".join(iter_text(p, TEXT))) != "" or norm("".join(iter_text(p, DEL_TEXT))) != "":
            return False
        if has_protected_content(p):
            return False
        if paragraph_mark_has_revision(p):
            return False
        return True

    def walk(parent: ET.Element) -> None:
        nonlocal changed
        # A table cell must retain at least one paragraph; never remove its last one.
        is_cell = parent.tag == W + "tc"
        for child in list(parent):
            if child.tag == W + "p":
                if is_empty_list_paragraph(child):
                    if is_cell and sum(1 for c in parent if c.tag == W + "p") <= 1:
                        continue
                    parent.remove(child)
                    changed += 1
            else:
                walk(child)

    walk(root)
    return changed


def _last_ins_last_run(p: ET.Element) -> tuple[ET.Element, ET.Element, str] | None:
    """Return (ins_elem, last_run, text) for the last <w:r> of p's last top-level <w:ins>."""
    last_ins = None
    for child in p:
        if child.tag == INS:
            last_ins = child
    if last_ins is None or has_protected_content(last_ins):
        return None
    last_run = None
    last_text = ""
    for r in last_ins:
        if r.tag != W + "r":
            continue
        text = "".join(t.text or "" for t in r.iter(TEXT))
        if text:
            last_run = r
            last_text = text
    if last_run is None:
        return None
    return last_ins, last_run, last_text


def clean_trailing_heading_ins_del_noop(root: ET.Element) -> int:
    """Clean a no-op where an unchanged heading is shown as both inserted and deleted.

    Word Compare sometimes appends an unchanged heading's text as the LAST inserted run of
    a preceding (real body-edit) paragraph, while marking the heading's true location as a
    full deletion a few paragraphs later. Detected pattern (forward window of 3):

    - Para N: its last top-level <w:ins> ends with a run whose text is T
    - Para B (optional): deletion-only paragraphs — skip and keep scanning
    - Para C (heading): a heading paragraph whose ENTIRE deleted text equals T

    Fix: drop the trailing <w:ins> run T from para N (keeping the real insertion), and turn
    para C's deletion of T into plain unchanged text (preserving its pPr/numPr/bookmark).
    """
    changed = 0

    def walk(parent: ET.Element) -> None:
        nonlocal changed
        paras = [child for child in parent if child.tag == W + "p"]

        for idx, p in enumerate(paras):
            found = _last_ins_last_run(p)
            if found is None:
                continue
            ins_elem, ins_run, ins_text = found
            if not norm(ins_text):
                continue

            for fwd in range(1, 4):
                if idx + fwd >= len(paras):
                    break
                target = paras[idx + fwd]

                dels = [c for c in target if c.tag == DEL]
                if dels:
                    if is_heading_paragraph(target):
                        del_text = "".join(iter_text(target, DEL_TEXT))
                        if norm(del_text) == norm(ins_text):
                            # Drop the trailing ins run T from para N.
                            ins_elem.remove(ins_run)
                            if not any(
                                r.tag == W + "r" and r.find(TEXT) is not None
                                for r in ins_elem
                            ):
                                p.remove(ins_elem)
                            # Turn each del in target into unchanged text.
                            for del_elem in dels:
                                del_index = list(target).index(del_elem)
                                target[del_index: del_index + 1] = del_to_unchanged_runs(del_elem)
                            changed += 1
                            break
                    # Heading (or any) del that doesn't match: stop unless deletion-only.
                if not is_deletion_only_paragraph(target):
                    break

        for child in parent:
            if child.tag != W + "p":
                walk(child)

    walk(root)
    return changed


def build_parent_map(root: ET.Element) -> dict[int, ET.Element]:
    """Map id(element) -> parent for every descendant (ElementTree has no parent pointers)."""
    parents: dict[int, ET.Element] = {}
    for parent in root.iter():
        for child in parent:
            parents[id(child)] = parent
    return parents


def visible_text(p: ET.Element) -> str:
    """The text a reader sees: unchanged runs plus insertions, but not deleted runs."""
    return "".join(iter_text(p, TEXT))


def _is_contentless_paragraph(p: ET.Element) -> bool:
    """True if p holds no text, no content revision and nothing but (maybe) bookmarks.

    Deliberately ignores the paragraph *mark's* own revision — that is a separate question,
    answered by paragraph_mark_revision. See is_spacer_paragraph / is_rejectable_spacer_paragraph.
    """
    if p.tag != W + "p":
        return False
    if norm(visible_text(p)) != "" or norm("".join(iter_text(p, DEL_TEXT))) != "":
        return False
    mark_rPr = paragraph_mark_rPr(p)
    mark_revisions = {id(child) for child in mark_rPr} if mark_rPr is not None else set()
    for node in p.iter():
        if node.tag in {INS, DEL} and id(node) not in mark_revisions:
            return False
    blocking = PROTECTED_TAGS - BOOKMARK_TAGS
    if any(node.tag in blocking for node in p.iter()):
        return False
    return not has_relationship_reference(p)


def is_spacer_paragraph(p: ET.Element) -> bool:
    """A pure blank line with no tracked state at all.

    Word Compare leaves these behind when the two documents' block structures differ — the
    revised file's blank lines are adopted untracked, so they survive both Accept and Reject
    and can only be removed by deleting the <w:p>.
    """
    return _is_contentless_paragraph(p) and paragraph_mark_revision(p) is None


def is_rejectable_spacer_paragraph(p: ET.Element) -> bool:
    """A blank line that is either untracked or one the revision *inserted*.

    A converted revision's layout noise usually arrives as inserted blank paragraphs, so
    Compare marks the paragraph mark itself with <w:ins>. Deleting such a paragraph is
    exactly what Word's Reject does to it — nothing the author wrote is lost.

    A *deleted* paragraph mark is left alone: that one is the reviewer's call to accept, and
    removing it here would pre-empt them.
    """
    return _is_contentless_paragraph(p) and paragraph_mark_revision(p) in (None, INS)


def relocate_bookmarks(source: ET.Element, target: ET.Element | None) -> None:
    """Move source's bookmark anchors into target (right after its pPr) before deleting it."""
    marks = [child for child in source if child.tag in BOOKMARK_TAGS]
    for mark in marks:
        source.remove(mark)
    if target is None:
        return
    insert_at = 1 if len(target) and target[0].tag == W + "pPr" else 0
    for offset, mark in enumerate(marks):
        target.insert(insert_at + offset, mark)


def _following_paragraph(parent: ET.Element, p: ET.Element) -> ET.Element | None:
    siblings = list(parent)
    for candidate in siblings[siblings.index(p) + 1:]:
        if candidate.tag == W + "p":
            return candidate
    return None


def _paragraph_is_removable(parent: ET.Element, p: ET.Element) -> bool:
    """True if p can be deleted from parent without leaving invalid OOXML."""
    paragraph_count = sum(1 for child in parent if child.tag == W + "p")
    if parent.tag in {W + "tc", W + "body"} and paragraph_count <= 1:
        return False
    siblings = list(parent)
    index = siblings.index(p)
    previous = siblings[index - 1] if index > 0 else None
    following = siblings[index + 1] if index + 1 < len(siblings) else None
    # A paragraph between two tables keeps them apart; removing it merges them into one
    # table with a mismatched grid.
    if previous is not None and following is not None:
        if previous.tag == W + "tbl" and following.tag == W + "tbl":
            return False
    return True


def _page_size(sect: ET.Element) -> tuple[str | None, str | None, str | None] | None:
    """Return (w, h, orient) of a <w:sectPr>'s <w:pgSz>, or None if absent."""
    pgSz = sect.find(W + "pgSz")
    if pgSz is None:
        return None
    return (pgSz.get(W + "w"), pgSz.get(W + "h"), pgSz.get(W + "orient"))


def clean_cosmetic_section_breaks(root: ET.Element) -> int:
    """Remove mid-document section breaks that Word Compare silently injects.

    With `detect format changes false`, Word Compare adopts the revised document's section
    structure without tracking it, so section breaks that exist only in the revised file
    appear in the output as untracked, non-rejectable page breaks. We strip a mid-document
    <w:sectPr> only when its page size/orientation matches the document's final (body-level)
    section — i.e. it is a purely cosmetic page break, not a landscape/differently-sized
    section. The body-level <w:sectPr> (whole-document page setup) is never removed.

    The paragraph that *hosted* such a break is almost always an empty one the revised file
    used as a break carrier, and leaving it behind shows up as a stray blank line above the
    next heading. So the host paragraph is deleted too, but only when it is a pure spacer
    (is_rejectable_spacer_paragraph) and removing it is safe (_paragraph_is_removable): its
    bookmarks move to the following paragraph first. Anything with text or content is kept.
    """
    body = root.find(W + "body")
    if body is None:
        return 0
    body_sect = body.find(W + "sectPr")
    if body_sect is None:
        return 0
    body_pgsz = _page_size(body_sect)
    if body_pgsz is None:
        return 0

    parents = build_parent_map(body)
    changed = 0
    for p in list(body.iter(W + "p")):
        pPr = p.find(W + "pPr")
        if pPr is None:
            continue
        sect = pPr.find(W + "sectPr")
        if sect is None:
            continue
        if _page_size(sect) != body_pgsz:
            continue
        pPr.remove(sect)
        changed += 1

        parent = parents.get(id(p))
        if parent is None or not is_rejectable_spacer_paragraph(p):
            continue
        if not _paragraph_is_removable(parent, p):
            continue
        relocate_bookmarks(p, _following_paragraph(parent, p))
        parent.remove(p)
        changed += 1
    return changed


def clean_leading_empty_paragraphs(root: ET.Element) -> int:
    """Drop the blank paragraphs Word Compare stacks up above the document title.

    The original and the revised file each carry their own spacer above the title (in this
    template the original keeps the logo and its 16pt gap in a single paragraph). Compare
    adopts the revised file's spacers untracked while keeping the original's, so the title
    drifts down the page by a blank line or two.

    Every pure spacer between the last content-bearing paragraph and the first paragraph with
    text is removed, and the gap they provided is folded back onto that content paragraph as
    pPr/spacing/@w:after — the largest explicit `after` among the removed spacers, or
    TITLE_GAP_AFTER. The logo paragraph itself holds a <w:drawing>, so it is not contentless
    and bounds the scan instead of being deleted.
    """
    body = root.find(W + "body")
    if body is None:
        return 0

    children = list(body)
    first_text_index = None
    for index, child in enumerate(children):
        if child.tag != W + "p":
            return 0  # a table before any text: not the letterhead pattern
        if norm(visible_text(child)) or norm("".join(iter_text(child, DEL_TEXT))):
            first_text_index = index
            break
    if first_text_index is None:
        return 0

    start = first_text_index
    while start > 0 and is_rejectable_spacer_paragraph(children[start - 1]):
        start -= 1
    if start == first_text_index:
        return 0

    spacers = children[start:first_text_index]
    following = children[first_text_index]
    previous = children[start - 1] if start > 0 else None

    gaps = []
    for spacer in spacers:
        pPr = spacer.find(W + "pPr")
        spacing = pPr.find(W + "spacing") if pPr is not None else None
        value = spacing.get(W + "after") if spacing is not None else None
        if value and value.isdigit():
            gaps.append(int(value))
    gap = str(max(gaps)) if gaps else TITLE_GAP_AFTER

    for spacer in spacers:
        relocate_bookmarks(spacer, following)
        body.remove(spacer)

    if previous is not None:
        pPr = previous.find(W + "pPr")
        if pPr is None:
            pPr = ET.Element(W + "pPr")
            previous.insert(0, pPr)
        spacing = _get_or_insert_ordered(pPr, "spacing", PPR_ORDER)
        spacing.set(W + "after", gap)

    return len(spacers)


def _sole_text_revision(p: ET.Element) -> tuple[str, str, ET.Element] | None:
    """Return (kind, text, elem) when p's only visible content is one text revision."""
    payloads = text_revision_payloads(p)
    if len(payloads) != 1:
        return None
    kind, text, revision = payloads[0]
    if not norm(text):
        return None
    if has_visible_text_except_revision(p, revision):
        return None
    return kind, text, revision


def _is_empty_table(table: ET.Element) -> bool:
    """True if a <w:tbl> holds no text and nothing worth keeping.

    Inserted paragraph marks and inserted rows are tolerated — the stray table a split
    leaves behind is usually marked wholly inserted, and dropping it is what Reject does.
    A *deleted* mark or row is left alone, as is any content revision.
    """
    if norm(visible_text(table)) != "" or norm("".join(iter_text(table, DEL_TEXT))) != "":
        return False
    for p in table.iter(W + "p"):
        if not is_rejectable_spacer_paragraph(p):
            return False
    # Row/cell markers live outside any paragraph; an inserted row may go, a deleted one stays.
    for local in ("trPr", "tcPr", "tblPr", "tblPrEx"):
        for props in table.iter(W + local):
            if props.find(DEL) is not None:
                return False
    for node in table.iter():
        if node is not table and node.tag in PROTECTED_TAGS:
            return False
    return not has_relationship_reference(table)


def _drop_emptied_container(p: ET.Element, parents: dict[int, ET.Element]) -> int:
    """After a relocation, remove the paragraph left empty and its table if that empties it."""
    parent = parents.get(id(p))
    if parent is None or not is_rejectable_spacer_paragraph(p):
        return 0
    if any(node.tag in BOOKMARK_TAGS for node in p.iter()) or has_relationship_reference(p):
        return 0

    changed = 0
    if _paragraph_is_removable(parent, p):
        parent.remove(p)
        changed += 1

    table = parent
    while table is not None and table.tag != W + "tbl":
        table = parents.get(id(table))
    if table is not None and _is_empty_table(table):
        table_parent = parents.get(id(table))
        if table_parent is not None:
            table_parent.remove(table)
            changed += 1
    return changed


def clean_relocated_table_noops(root: ET.Element) -> int:
    """Undo a no-op that Word Compare spread across a table boundary.

    When the revised file splits a table (a converted document often drops a section break
    mid-table), Compare cannot align the two halves: it deletes a row's text from the
    original table and re-inserts the same text — often differing only in whitespace — into
    the new one, so the redline claims content moved when it never did.

    clean_cross_paragraph_noops only compares adjacent sibling paragraphs, so it never sees
    across the boundary. This rule matches a DEL-only paragraph against an INS-only paragraph
    within RELOCATION_WINDOW paragraphs of flattened document order, requiring the two to sit
    in *different* parents with at least one inside a table.

    Unlike the adjacent-sibling rule it keeps the DELETION's location — the original position
    is the true one — restoring it as unchanged text and dropping the insertion. If that
    leaves the insertion's paragraph empty it goes too, and if the table it lived in is then
    completely empty the whole <w:tbl> is removed.
    """
    body = root.find(W + "body")
    if body is None:
        return 0

    parents = build_parent_map(body)
    paragraphs = list(body.iter(W + "p"))

    def inside_table(p: ET.Element) -> bool:
        node = parents.get(id(p))
        while node is not None:
            if node.tag == W + "tbl":
                return True
            node = parents.get(id(node))
        return False

    changed = 0
    for index, first in enumerate(paragraphs):
        found = _sole_text_revision(first)
        if found is None or found[0] != DEL:
            continue
        _, del_text, del_revision = found

        for offset in range(1, RELOCATION_WINDOW + 1):
            if index + offset >= len(paragraphs):
                break
            second = paragraphs[index + offset]
            if parents.get(id(second)) is parents.get(id(first)):
                continue
            if not (inside_table(first) or inside_table(second)):
                continue
            match = _sole_text_revision(second)
            if match is None or match[0] != INS:
                continue
            if not same_text_or_spacing_only(del_text, match[1]):
                continue

            replace_descendant(first, del_revision, del_to_unchanged_runs(del_revision))
            replace_descendant(second, match[2], [])
            changed += 1 + _drop_emptied_container(second, parents)
            break

    return changed


# Punctuation that ends a sentence or a clause — a paragraph ending in one of these was not
# split mid-sentence, so it is never a merge candidate.
SENTENCE_BREAK_CHARS = frozenset(".?!:;,")

# Content that must not be merged across paragraph boundaries. Hyperlinks and bookmarks are
# fine (they simply move with the runs); pictures, fields, tables and comment anchors are not.
MERGE_BLOCKING_TAGS = PROTECTED_TAGS - HYPERLINK_TAGS - BOOKMARK_TAGS

# Blank paragraphs tolerated between the two halves of a split sentence. A converted
# revision leaves one where the page furniture sat and another carrying its section break,
# so both must be crossable; the lower-case-start test is what actually gates the merge.
MERGE_MAX_SPACERS = 2


def _can_continue_sentence(p: ET.Element) -> bool:
    if is_heading_paragraph(p) or paragraph_style(p) in NON_BODY_STYLES:
        return False
    pPr = p.find(W + "pPr")
    if pPr is not None:
        if pPr.find(W + "numPr") is not None or pPr.find(W + "sectPr") is not None:
            return False
    # An inserted paragraph mark is the conversion's own break — merging rejects it. A
    # deleted mark already tells the reviewer these two join, so leave that to them.
    if paragraph_mark_revision(p) == DEL:
        return False
    return not any(node.tag in MERGE_BLOCKING_TAGS for node in p.iter())


def _merge_paragraphs(first: ET.Element, second: ET.Element) -> None:
    """Append second's content to first, inserting a separating space where needed."""
    last_text = None
    for node in first.iter(TEXT):
        if node.text:
            last_text = node
    if last_text is not None and not last_text.text[-1:].isspace():
        if not visible_text(second)[:1].isspace():
            last_text.text += " "
            last_text.set(XML_SPACE, "preserve")
    for child in list(second):
        if child.tag == W + "pPr":
            continue
        first.append(copy.deepcopy(child))


def clean_conversion_paragraph_splits(root: ET.Element) -> int:
    """Rejoin a sentence a converted revision split across two paragraphs.

    A PDF-derived revision breaks running text wherever the page did — around a flattened
    header/footer, for example — so one sentence arrives as two paragraphs. Word Compare
    adopts the split untracked, leaving a paragraph that stops mid-sentence and a next one
    that starts lower-case.

    Merge only on that exact evidence: the first paragraph's visible text ends in a letter
    (no sentence or clause terminator) and the second's starts with a lower-case letter.
    Both must be plain body paragraphs (no heading/Title style, no list numbering, no tracked
    paragraph mark, no section break, no pictures/fields/tables) and direct children of
    <w:body>, which also rules out anything inside a table. Up to MERGE_MAX_SPACERS blank
    paragraphs may sit between them; they are removed with the merge.
    """
    body = root.find(W + "body")
    if body is None:
        return 0

    changed = 0
    merged = True
    while merged:
        merged = False
        children = list(body)
        for index, first in enumerate(children):
            if first.tag != W + "p" or not _can_continue_sentence(first):
                continue
            text = norm(visible_text(first))
            if not text or not text[-1].isalpha():
                continue

            spacers: list[ET.Element] = []
            second = None
            for candidate in children[index + 1:]:
                if candidate.tag != W + "p":
                    break
                if is_rejectable_spacer_paragraph(candidate) and not has_protected_content(candidate):
                    if len(spacers) >= MERGE_MAX_SPACERS:
                        break
                    spacers.append(candidate)
                    continue
                second = candidate
                break

            if second is None or not _can_continue_sentence(second):
                continue
            next_text = norm(visible_text(second))
            if not next_text[:1].isalpha() or not next_text[:1].islower():
                continue

            _merge_paragraphs(first, second)
            for spacer in spacers:
                body.remove(spacer)
            body.remove(second)
            changed += 1 + len(spacers)
            merged = True
            break

    return changed


DOCUMENT_PART = "word/document.xml"


def clean_header_footer_relocations(
    doc_root: ET.Element, hf_roots: dict[str, ET.Element]
) -> dict[str, int]:
    """Restore header/footer text that a converted revision flattened into the body.

    A PDF-derived revision has no header/footer parts at all — the page furniture is baked
    into the body text. Word Compare therefore reports the original's footer as deleted while
    the same string shows up as an insertion in the middle of the body, which is how a footer
    ends up printed mid-paragraph in the redline.

    Match a header/footer paragraph whose only revision is a DEL against a body paragraph
    whose only revision is an INS of the same text (whitespace-insensitive): restore the
    header/footer del in place and delete the body insertion. Returns the change count per
    part so the caller knows which parts to rewrite.
    """
    body = doc_root.find(W + "body")
    if body is None:
        return {}

    parents = build_parent_map(body)
    candidates = list(body.iter(W + "p"))
    counts: dict[str, int] = {}

    for name, hf_root in hf_roots.items():
        for hf_p in list(hf_root.iter(W + "p")):
            found = _sole_text_revision(hf_p)
            if found is None or found[0] != DEL:
                continue
            _, del_text, del_revision = found

            match = None
            for body_p in candidates:
                payload = _sole_text_revision(body_p)
                if payload is None or payload[0] != INS:
                    continue
                if same_text_or_spacing_only(del_text, payload[1]):
                    match = (body_p, payload[2])
                    break
            if match is None:
                continue

            body_p, ins_revision = match
            replace_descendant(hf_p, del_revision, del_to_unchanged_runs(del_revision))
            replace_descendant(body_p, ins_revision, [])
            counts[name] = counts.get(name, 0) + 1
            counts[DOCUMENT_PART] = (
                counts.get(DOCUMENT_PART, 0) + 1 + _drop_emptied_container(body_p, parents)
            )
            candidates = [p for p in candidates if p is not body_p]

    return counts


def _has_image(elem: ET.Element) -> bool:
    return any(node.tag in IMAGE_TAGS for node in elem.iter())


def _has_ole(elem: ET.Element) -> bool:
    """Embedded OLE objects (spreadsheets, equations) are not logos — never touch them."""
    return any(node.tag in OLE_TAGS for node in elem.iter())


def _is_image_only_revision(elem: ET.Element) -> bool:
    """True if `elem` is a <w:ins>/<w:del> whose entire content is an image."""
    if elem.tag not in {INS, DEL}:
        return False
    if not _has_image(elem) or _has_ole(elem):
        return False
    return revision_text(elem).strip() == ""


def _image_runs(revision: ET.Element) -> list[ET.Element]:
    """The revision's direct <w:r> children that carry an image.

    Only direct run children are handled. An image nested deeper (inside a hyperlink, say)
    is left alone rather than half-resolved — and, importantly, reporting it as unchanged is
    what keeps clean_root's fixed-point loop terminating.
    """
    return [
        run for run in revision
        if run.tag == W + "r" and any(node.tag in IMAGE_TAGS for node in run.iter())
    ]


def _resolve_image_runs(revision: ET.Element) -> list[ET.Element]:
    """Strip the image content out of a mixed revision, returning runs to place before it.

    A revision that mixes an image with real text is split rather than skipped: the image
    is resolved (kept unmarked for an insertion, dropped for a deletion) while the text
    stays tracked. Returns the unmarked runs the caller must insert ahead of `revision`.
    """
    keep_unmarked: list[ET.Element] = []
    is_insertion = revision.tag == INS

    for run in _image_runs(revision):
        run_text = "".join(
            node.text or "" for node in run.iter() if node.tag in {TEXT, DEL_TEXT}
        )
        if run_text.strip() == "":
            # Whole run is just the image.
            revision.remove(run)
            if is_insertion:
                keep_unmarked.append(copy.deepcopy(run))
        else:
            # Image and text share one run: lift the image out, leave the text tracked.
            images = [node for node in run if node.tag in IMAGE_TAGS]
            if is_insertion:
                carrier = ET.Element(W + "r")
                rPr = run.find(W + "rPr")
                if rPr is not None:
                    carrier.append(copy.deepcopy(rPr))
                for image in images:
                    carrier.append(copy.deepcopy(image))
                keep_unmarked.append(carrier)
            for image in images:
                run.remove(image)

    return keep_unmarked


def clean_image_revisions(root: ET.Element) -> int:
    """Never show an image (the logo) as a tracked change.

    Word Compare marks a logo as an insert+delete pair of the picture whenever the two
    documents reference it through different relationships — the image is identical, so
    the redline is always false. Accept image insertions (unwrap to a normal, unchanged
    run) and drop image deletions; a delete+insert pair collapses to one unmarked image.

    A revision that mixes an image with real text is **split**, not skipped: the image is
    resolved and the text stays tracked. That is what lets the logo be resolved when Word
    sweeps a stray character into the same <w:del> as the picture. The consequence, accepted
    deliberately: a genuine "figure and its caption deleted" edit shows the caption struck
    through while the figure simply disappears. Embedded OLE objects are never touched.
    """
    changed = 0

    def walk(parent: ET.Element) -> None:
        nonlocal changed
        i = 0
        while i < len(parent):
            child = parent[i]

            if _is_image_only_revision(child):
                if child.tag == INS:
                    replacement = accepted_children(child)
                    parent[i: i + 1] = replacement
                    i += len(replacement)
                else:  # DEL
                    parent.remove(child)
                changed += 1
                continue

            # A revision mixing an image with text: resolve only the image runs. Guarded on
            # _image_runs being non-empty so a no-op never reports a change (which would
            # spin clean_root's fixed-point loop forever).
            if (
                child.tag in {INS, DEL}
                and not _has_ole(child)
                and _image_runs(child)
            ):
                unmarked = _resolve_image_runs(child)
                for offset, run in enumerate(unmarked):
                    parent.insert(i + offset, run)
                i += len(unmarked)
                changed += 1
                # The revision may now be empty (all its runs were images).
                if len(parent[i]) == 0:
                    parent.remove(parent[i])
                    continue
                i += 1
                continue

            walk(child)
            i += 1

    walk(root)
    return changed


def _get_or_insert_ordered(parent: ET.Element, local: str, order: tuple[str, ...]) -> ET.Element:
    """Return parent's child <w:{local}>, inserting a new one in schema order if absent."""
    tag = W + local
    existing = parent.find(tag)
    if existing is not None:
        return existing
    elem = ET.Element(tag)
    rank = order.index(local)
    insert_at = len(parent)
    for i, child in enumerate(parent):
        name = child.tag.split("}")[-1]
        if name in order and order.index(name) > rank:
            insert_at = i
            break
    parent.insert(insert_at, elem)
    return elem


def _set_half_point_size(rPr: ET.Element) -> int:
    """Set w:sz and w:szCs to the body font size on an rPr; return 1 if anything changed."""
    changed = 0
    for local in ("sz", "szCs"):
        el = _get_or_insert_ordered(rPr, local, RPR_ORDER)
        if el.get(W + "val") != BODY_FONT_HALF_PT:
            el.set(W + "val", BODY_FONT_HALF_PT)
            changed = 1
    return changed


def normalize_body_formatting(root: ET.Element) -> int:
    """Force body paragraphs to 10pt / 1.15 line spacing (silent direct formatting).

    Applies to normal body paragraphs and table paragraphs; skips headings and the
    Title/Subtitle so their sizes/spacing are preserved. Only touches the main document
    (there is no <w:body> in header/footer parts). Idempotent — re-running makes no change.
    """
    body = root.find(W + "body")
    if body is None:
        return 0

    changed = 0
    for p in body.iter(W + "p"):
        style = paragraph_style(p)
        if is_heading_paragraph(p) or (style in NON_BODY_STYLES):
            continue

        pPr = p.find(W + "pPr")
        if pPr is None:
            pPr = ET.Element(W + "pPr")
            p.insert(0, pPr)  # pPr is always the first child of a paragraph
        # Line spacing 1.15
        spacing = _get_or_insert_ordered(pPr, "spacing", PPR_ORDER)
        if spacing.get(W + "line") != BODY_LINE or spacing.get(W + "lineRule") != "auto":
            spacing.set(W + "line", BODY_LINE)
            spacing.set(W + "lineRule", "auto")
            changed = 1
        # Paragraph-mark font size (empty body lines render at 10pt too)
        mark_rPr = _get_or_insert_ordered(pPr, "rPr", PPR_ORDER)
        changed |= _set_half_point_size(mark_rPr)
        # Every run's font size
        for r in p.iter(W + "r"):
            run_rPr = r.find(W + "rPr")
            if run_rPr is None:
                run_rPr = ET.Element(W + "rPr")
                r.insert(0, run_rPr)  # rPr is always the first child of a run
            changed |= _set_half_point_size(run_rPr)

    return 1 if changed else 0


def clean_root(root: ET.Element) -> int:
    """Run every artifact rule over one XML part until nothing more changes.

    Order matters in one place: clean_relocated_table_noops runs before
    clean_cosmetic_section_breaks so a stray table it empties is gone before the
    between-two-tables guard on the break's host paragraph is evaluated.
    """
    total_changed = 0
    while True:
        changed = (
            clean_parent(root)
            + clean_cross_paragraph_noops(root)
            + clean_misplaced_heading_insertions(root)
            + clean_paragraph_boundary_noop(root)
            + clean_heading_full_ins_noop(root)
            + clean_misplaced_deletions_after_heading(root)
            + clean_trailing_heading_ins_del_noop(root)
            + clean_empty_list_paragraphs(root)
            + clean_image_revisions(root)
            + clean_relocated_table_noops(root)
            + clean_cosmetic_section_breaks(root)
            + clean_leading_empty_paragraphs(root)
            + clean_conversion_paragraph_splits(root)
        )
        if changed == 0:
            return total_changed
        total_changed += changed


def clean_document_xml(xml_bytes: bytes) -> tuple[bytes, int]:
    """Clean a single XML part in isolation (no cross-part rules)."""
    root = ET.fromstring(xml_bytes)
    changed = clean_root(root)
    changed += normalize_body_formatting(root)
    # Last: removing revisions orphans the prefixes they used (e.g. w16du:dateUtc), and
    # pruning before the loop would leave them behind for a second run to notice.
    changed += clean_ignorable_namespaces(root)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True), changed


def clean_docx_parts(parts: dict[str, bytes]) -> dict[str, tuple[bytes, int]]:
    """Clean every cleanable part together, so cross-part artifacts can be resolved.

    Header/footer content that the revised document flattened into the body can only be
    recognised with the body and the header/footer parts in scope at once, so all parts are
    parsed first, cleaned individually, then reconciled and cleaned again to a fixed point.
    """
    roots: dict[str, ET.Element] = {}
    counts: dict[str, int] = {}
    for name, data in parts.items():
        root = ET.fromstring(data)
        roots[name] = root
        counts[name] = clean_root(root)

    doc_root = roots.get(DOCUMENT_PART)
    if doc_root is not None:
        hf_roots = {name: root for name, root in roots.items() if name != DOCUMENT_PART}
        cross = clean_header_footer_relocations(doc_root, hf_roots)
        for name, count in cross.items():
            counts[name] = counts.get(name, 0) + count
        if cross:
            for name, root in roots.items():
                counts[name] += clean_root(root)

    result: dict[str, tuple[bytes, int]] = {}
    for name, root in roots.items():
        counts[name] += normalize_body_formatting(root)
        # Last, once every revision that referenced a prefix is gone — see clean_document_xml.
        counts[name] += clean_ignorable_namespaces(root)
        result[name] = (
            ET.tostring(root, encoding="utf-8", xml_declaration=True),
            counts[name],
        )
    return result


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


# Parts that carry body-style content and can hold a logo revision: the main document plus
# every header/footer (logos live in page headers for some templates).
CLEANABLE_PART_RE = re.compile(r"^word/(document|header\d+|footer\d+)\.xml$")


def rewrite_docx(path: Path) -> int:
    with ZipFile(path, "r") as zin:
        infos = zin.infolist()
        parts = {
            info.filename: zin.read(info.filename)
            for info in infos
            if CLEANABLE_PART_RE.match(info.filename)
        }
        cleaned = clean_docx_parts(parts)
        cleaned_parts = {name: data for name, (data, count) in cleaned.items() if count}
        changed = sum(count for _, count in cleaned.values())

        if changed == 0:
            return 0

        fd, tmp_name = tempfile.mkstemp(suffix=".docx", dir=str(path.parent))
        os.close(fd)
        tmp_path = Path(tmp_name)
        try:
            with ZipFile(tmp_path, "w", ZIP_DEFLATED) as zout:
                for info in infos:
                    data = cleaned_parts.get(info.filename) or zin.read(info.filename)
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
