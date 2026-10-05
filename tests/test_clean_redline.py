import unittest
import xml.etree.ElementTree as ET

import clean_redline


W = clean_redline.W
R = clean_redline.R
W14 = clean_redline.W14
V = f"{{{clean_redline.NAMESPACES['v']}}}"
O = f"{{{clean_redline.NAMESPACES['o']}}}"


def clean_fragment(body: str) -> ET.Element:
    xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:document
    xmlns:w="{clean_redline.NAMESPACES['w']}"
    xmlns:r="{clean_redline.NAMESPACES['r']}"
    xmlns:v="{clean_redline.NAMESPACES['v']}"
    xmlns:o="{clean_redline.NAMESPACES['o']}"
    xmlns:w14="{clean_redline.NAMESPACES['w14']}">
  <w:body>{body}</w:body>
</w:document>"""
    cleaned, _ = clean_redline.clean_document_xml(xml.encode("utf-8"))
    return ET.fromstring(cleaned)


class CleanRedlineTests(unittest.TestCase):
    def test_removes_empty_text_only_revision(self):
        root = clean_fragment("<w:p><w:ins w:id=\"1\"/></w:p>")

        self.assertEqual(root.findall(f".//{W}ins"), [])

    def test_accepts_whitespace_only_insert(self):
        root = clean_fragment(
            "<w:p><w:ins w:id=\"1\"><w:r><w:t>   </w:t></w:r></w:ins></w:p>"
        )

        self.assertEqual(root.findall(f".//{W}ins"), [])
        self.assertEqual(root.find(f".//{W}t").text, "   ")

    def test_collapses_identical_text_revision_pair(self):
        root = clean_fragment(
            "<w:p>"
            "<w:del w:id=\"1\"><w:r><w:delText>same</w:delText></w:r></w:del>"
            "<w:ins w:id=\"2\"><w:r><w:t>same</w:t></w:r></w:ins>"
            "</w:p>"
        )

        self.assertEqual(root.findall(f".//{W}del"), [])
        self.assertEqual(root.findall(f".//{W}ins"), [])
        self.assertEqual(root.find(f".//{W}t").text, "same")

    def test_collapses_spacing_only_revision_pair_to_inserted_text(self):
        root = clean_fragment(
            "<w:p><w:r><w:t>1.</w:t></w:r>"
            "<w:ins w:id=\"1\"><w:r><w:t xml:space=\"preserve\">4 </w:t></w:r><w:r><w:t>Independent</w:t></w:r></w:ins>"
            "<w:del w:id=\"2\"><w:r><w:delText>4Independent</w:delText></w:r></w:del>"
            "<w:r><w:t> Contractor Status.</w:t></w:r></w:p>"
        )

        self.assertEqual(root.findall(f".//{W}del"), [])
        self.assertEqual(root.findall(f".//{W}ins"), [])
        self.assertEqual("".join(root.itertext()).strip(), "1.4 Independent Contractor Status.")

    def test_accepts_linked_drawing_insertion(self):
        # A linked image (logo) insertion is accepted to an unchanged run — never shown
        # as a tracked change (per the "images never show as changed" rule).
        root = clean_fragment(
            "<w:p><w:ins w:id=\"1\"><w:r><w:drawing r:id=\"rId5\"/></w:r></w:ins></w:p>"
        )

        self.assertEqual(root.findall(f".//{W}ins"), [])
        self.assertEqual(root.find(f".//{W}drawing").attrib[f"{R}id"], "rId5")

    def test_removes_empty_revision_with_unlinked_drawing(self):
        root = clean_fragment(
            "<w:p><w:ins w:id=\"1\"><w:r><w:drawing/></w:r></w:ins></w:p>"
        )

        self.assertEqual(root.findall(f".//{W}ins"), [])
        self.assertEqual(root.findall(f".//{W}drawing"), [])

    def test_preserves_empty_revision_with_vml_ole_object(self):
        root = clean_fragment(
            "<w:p><w:del w:id=\"1\"><w:r><w:object>"
            "<v:shape><v:imagedata r:id=\"rId6\"/></v:shape>"
            "<o:OLEObject r:id=\"rId7\"/>"
            "</w:object></w:r></w:del></w:p>"
        )

        self.assertEqual(len(root.findall(f".//{W}del")), 1)
        self.assertIsNotNone(root.find(f".//{W}object"))
        self.assertEqual(root.find(f".//{V}imagedata").attrib[f"{R}id"], "rId6")
        self.assertEqual(root.find(f".//{O}OLEObject").attrib[f"{R}id"], "rId7")

    def test_preserves_comment_anchor_revision(self):
        root = clean_fragment(
            "<w:p><w:ins w:id=\"1\">"
            "<w:commentRangeStart w:id=\"4\"/>"
            "<w:r><w:commentReference w:id=\"4\"/></w:r>"
            "<w:commentRangeEnd w:id=\"4\"/>"
            "</w:ins></w:p>"
        )

        self.assertEqual(len(root.findall(f".//{W}ins")), 1)
        self.assertEqual(len(root.findall(f".//{W}commentRangeStart")), 1)
        self.assertEqual(len(root.findall(f".//{W}commentRangeEnd")), 1)
        self.assertEqual(len(root.findall(f".//{W}commentReference")), 1)

    def test_unwraps_standalone_checkbox_content_control(self):
        root = clean_fragment(
            "<w:p><w:sdt><w:sdtPr><w14:checkbox/></w:sdtPr>"
            "<w:sdtContent><w:r><w:t>☐</w:t></w:r></w:sdtContent>"
            "</w:sdt><w:r><w:t xml:space=\"preserve\"> Recipient</w:t></w:r></w:p>"
        )

        self.assertEqual(root.findall(f".//{W}sdt"), [])
        self.assertEqual("".join(root.itertext()).strip(), "☐ Recipient")

    def test_removes_checkbox_content_control_when_next_run_has_checkbox(self):
        root = clean_fragment(
            "<w:p><w:sdt><w:sdtPr><w14:checkbox/></w:sdtPr>"
            "<w:sdtContent><w:r><w:t>☐</w:t></w:r></w:sdtContent>"
            "</w:sdt><w:r><w:t>☒ Transferor</w:t></w:r></w:p>"
        )

        self.assertEqual(root.findall(f".//{W}sdt"), [])
        self.assertEqual("".join(root.itertext()).strip(), "☒ Transferor")

    def test_collapses_cross_paragraph_identical_delete_insert(self):
        root = clean_fragment(
            "<w:p><w:r><w:t>set forth in </w:t></w:r>"
            "<w:del w:id=\"1\"><w:r><w:delText>this Agreement;</w:delText></w:r>"
            "<w:r><w:delText xml:space=\"preserve\"> and</w:delText></w:r></w:del></w:p>"
            "<w:p><w:ins w:id=\"2\"><w:r><w:t xml:space=\"preserve\">this Agreement; and </w:t></w:r></w:ins></w:p>"
        )

        self.assertEqual(root.findall(f".//{W}del"), [])
        self.assertEqual(root.findall(f".//{W}ins"), [])
        self.assertEqual("".join(root.itertext()).strip(), "set forth in this Agreement; and")

    def test_moves_misplaced_heading_insertion_to_empty_heading(self):
        root = clean_fragment(
            "<w:p>"
            "<w:ins w:id=\"1\"><w:r><w:t>Reimbursable Expenses</w:t></w:r></w:ins>"
            "<w:del w:id=\"2\"><w:r><w:delText>old text</w:delText></w:r></w:del>"
            "</w:p>"
            "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr></w:p>"
        )

        body_paras = root.findall(f".//{W}p")
        body_p = body_paras[0]
        heading_p = body_paras[1]

        # Insertion must have moved to the heading
        self.assertIsNotNone(heading_p.find(f".//{W}ins"))
        self.assertEqual(
            "".join(heading_p.find(f".//{W}ins").itertext()),
            "Reimbursable Expenses",
        )
        # Body paragraph must no longer contain the insertion
        self.assertIsNone(body_p.find(f".//{W}ins"))
        # Deletion in body paragraph must remain
        self.assertIsNotNone(body_p.find(f".//{W}del"))

    def test_moves_misplaced_heading_insertion_skipping_deletion_only_paragraph(self):
        root = clean_fragment(
            "<w:p>"
            "<w:ins w:id=\"1\"><w:r><w:t>Section Title</w:t></w:r></w:ins>"
            "</w:p>"
            "<w:p><w:del w:id=\"2\"><w:r><w:delText>deleted only</w:delText></w:r></w:del></w:p>"
            "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr></w:p>"
        )

        # The emptied source paragraph is now a leading blank and gets removed, so find the
        # heading by style rather than by position.
        heading_p = next(
            p for p in root.findall(f".//{W}p")
            if clean_redline.paragraph_style(p) == "Heading1"
        )

        self.assertIsNotNone(heading_p.find(f".//{W}ins"))
        self.assertEqual(
            "".join(heading_p.find(f".//{W}ins").itertext()),
            "Section Title",
        )

    def test_cleans_cross_paragraph_boundary_noop_in_heading(self):
        root = clean_fragment(
            "<w:p>"
            "<w:r><w:t xml:space=\"preserve\">body text </w:t></w:r>"
            "<w:del w:id=\"1\"><w:r><w:delText>Representations</w:delText></w:r></w:del>"
            "</w:p>"
            "<w:p>"
            "<w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr>"
            "<w:ins w:id=\"2\"><w:r><w:t>Representations</w:t></w:r></w:ins>"
            "<w:r><w:t xml:space=\"preserve\"> and Warranties</w:t></w:r>"
            "</w:p>"
        )

        paras = root.findall(f".//{W}p")
        body_p = paras[0]
        heading_p = paras[1]

        # Deletion and insertion must both be gone
        self.assertEqual(body_p.findall(f".//{W}del"), [])
        self.assertEqual(heading_p.findall(f".//{W}ins"), [])
        # Heading must contain the full text as plain runs
        heading_text = "".join(heading_p.itertext())
        self.assertIn("Representations", heading_text)
        self.assertIn("and Warranties", heading_text)

    def test_boundary_noop_preserves_partial_del_when_last_run_matches(self):
        """When only the last run of a multi-run del matches the next-para ins, only that run is removed."""
        root = clean_fragment(
            "<w:p>"
            "<w:del w:id=\"1\">"
            "<w:r><w:delText>nature</w:delText></w:r>"
            "<w:r><w:delText>Representations</w:delText></w:r>"
            "</w:del>"
            "</w:p>"
            "<w:p>"
            "<w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr>"
            "<w:ins w:id=\"2\"><w:r><w:t>Representations</w:t></w:r></w:ins>"
            "<w:r><w:t xml:space=\"preserve\"> and Warranties</w:t></w:r>"
            "</w:p>"
        )

        paras = root.findall(f".//{W}p")
        body_p = paras[0]
        heading_p = paras[1]

        # Del must still exist (for "nature") but "Representations" run must be gone
        remaining_dels = body_p.findall(f".//{W}del")
        self.assertEqual(len(remaining_dels), 1)
        del_texts = [t.text for t in remaining_dels[0].findall(f".//{W}delText") if t.text]
        self.assertIn("nature", del_texts)
        self.assertNotIn("Representations", del_texts)

        # Heading: no insertion, full text present
        self.assertEqual(heading_p.findall(f".//{W}ins"), [])
        self.assertIn("Representations", "".join(heading_p.itertext()))


    def test_heading_with_only_ins_accepts_when_del_nearby_matches(self):
        """Heading has only <w:ins>T</w:ins>; a deletion-only para intervenes; next para del ends with T."""
        root = clean_fragment(
            "<w:p>"
            "<w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr>"
            "<w:ins w:id=\"1\"><w:r><w:t>Section Title</w:t></w:r></w:ins>"
            "</w:p>"
            "<w:p><w:del w:id=\"2\"><w:r><w:delText>old body text</w:delText></w:r></w:del></w:p>"
            "<w:p>"
            "<w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr>"
            "<w:del w:id=\"3\">"
            "<w:r><w:delText>other content </w:delText></w:r>"
            "<w:r><w:delText>Section Title</w:delText></w:r>"
            "</w:del>"
            "</w:p>"
        )
        paras = root.findall(f".//{W}p")
        heading_p = paras[0]
        del_p = paras[2]

        # Heading: ins gone, "Section Title" present as plain text
        self.assertIsNone(heading_p.find(f".//{W}ins"))
        self.assertIn("Section Title", "".join(heading_p.itertext()))

        # Del paragraph: "Section Title" run removed, "other content" remains
        del_texts = [t.text for t in del_p.findall(f".//{W}delText") if t.text]
        self.assertNotIn("Section Title", del_texts)
        self.assertTrue(any("other content" in t for t in del_texts))

    def test_inline_ins_before_del_suffix_match_accepts_ins_and_trims_del(self):
        """In same para, <w:ins>T</w:ins> before <w:del>...T</w:del>; last del run T removed."""
        root = clean_fragment(
            "<w:p>"
            "<w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr>"
            "<w:ins w:id=\"1\"><w:r><w:t>Word</w:t></w:r></w:ins>"
            "<w:del w:id=\"2\">"
            "<w:r><w:delText>prefix</w:delText></w:r>"
            "<w:r><w:delText>Word</w:delText></w:r>"
            "</w:del>"
            "<w:r><w:t xml:space=\"preserve\"> rest</w:t></w:r>"
            "</w:p>"
        )
        p = root.find(f".//{W}p")

        # Ins must be accepted
        self.assertIsNone(p.find(f".//{W}ins"))
        # "Word" run removed from del, "prefix" stays
        del_texts = [t.text for t in p.findall(f".//{W}delText") if t.text]
        self.assertIn("prefix", del_texts)
        self.assertNotIn("Word", del_texts)
        # "Word" and "rest" both present as accepted text
        self.assertIn("Word", "".join(p.itertext()))
        self.assertIn("rest", "".join(p.itertext()))


    # ── clean_misplaced_deletions_after_heading ───────────────────────────────

    def test_moves_misplaced_deletion_before_clean_heading(self):
        """Del-only para after a clean heading moves to before the heading."""
        root = clean_fragment(
            "<w:p>"
            "<w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr>"
            "<w:r><w:t>Section Title</w:t></w:r>"
            "</w:p>"
            "<w:p><w:del w:id=\"1\"><w:r><w:delText>old section 2 text</w:delText></w:r></w:del></w:p>"
            "<w:p><w:r><w:t>New section body.</w:t></w:r></w:p>"
        )
        paras = root.findall(f".//{W}p")
        # Deletion moved before heading
        self.assertIsNotNone(paras[0].find(f".//{W}del"))
        self.assertIsNone(paras[0].find(f".//{W}ins"))
        # Heading is now second
        pPr = paras[1].find(f"{W}pPr")
        pStyle = pPr.find(f"{W}pStyle") if pPr is not None else None
        self.assertEqual(pStyle.get(f"{W}val") if pStyle is not None else None, "Heading1")
        self.assertIn("Section Title", "".join(paras[1].itertext()))
        # Plain body is still third
        self.assertIn("New section body.", "".join(paras[2].itertext()))

    def test_moves_multiple_misplaced_deletions_before_heading(self):
        """Multiple contiguous del-only paras after a clean heading all move before it."""
        root = clean_fragment(
            "<w:p>"
            "<w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr>"
            "<w:r><w:t>Section Title</w:t></w:r>"
            "</w:p>"
            "<w:p><w:del w:id=\"1\"><w:r><w:delText>old text A</w:delText></w:r></w:del></w:p>"
            "<w:p><w:del w:id=\"2\"><w:r><w:delText>old text B</w:delText></w:r></w:del></w:p>"
            "<w:p><w:r><w:t>New section body.</w:t></w:r></w:p>"
        )
        paras = root.findall(f".//{W}p")
        # Both deletions moved before heading (order preserved)
        del_texts_0 = [t.text for t in paras[0].findall(f".//{W}delText") if t.text]
        del_texts_1 = [t.text for t in paras[1].findall(f".//{W}delText") if t.text]
        self.assertIn("old text A", del_texts_0)
        self.assertIn("old text B", del_texts_1)
        # Heading is now third
        self.assertIn("Section Title", "".join(paras[2].itertext()))

    def test_strips_heading_style_from_moved_deletion_paragraph(self):
        """A deletion-only para with heading style loses that style when moved before the heading."""
        root = clean_fragment(
            "<w:p>"
            "<w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr>"
            "<w:r><w:t>Section Title</w:t></w:r>"
            "</w:p>"
            "<w:p>"
            "<w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr>"
            "<w:del w:id=\"1\"><w:r><w:delText>old body text that Word Compare styled wrong</w:delText></w:r></w:del>"
            "</w:p>"
            "<w:p><w:r><w:t>New section body.</w:t></w:r></w:p>"
        )
        paras = root.findall(f".//{W}p")
        # Moved deletion para (now first) must NOT have heading style
        pPr = paras[0].find(f"{W}pPr")
        pStyle = pPr.find(f"{W}pStyle") if pPr is not None else None
        self.assertIsNone(pStyle, "heading style must be stripped from moved deletion para")
        # Actual heading is now second
        self.assertIn("Section Title", "".join(paras[1].itertext()))

    def test_does_not_move_deletion_when_heading_has_tracked_changes(self):
        """Del-only para after a heading-under-revision must NOT be moved."""
        root = clean_fragment(
            "<w:p>"
            "<w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr>"
            "<w:ins w:id=\"1\"><w:r><w:t>Section Title</w:t></w:r></w:ins>"
            "</w:p>"
            "<w:p><w:del w:id=\"2\"><w:r><w:delText>old text</w:delText></w:r></w:del></w:p>"
            "<w:p><w:r><w:t>New section body.</w:t></w:r></w:p>"
        )
        paras = root.findall(f".//{W}p")
        # The heading's ins may have been cleaned by clean_heading_full_ins_noop since
        # there's no nearby del with matching text — so we just check that the deletion
        # paragraph is NOT moved before the heading text.
        heading_idx = next(
            i for i, p in enumerate(paras)
            if "Section Title" in "".join(p.itertext())
        )
        del_para_idx = next(
            i for i, p in enumerate(paras)
            if p.find(f".//{W}del") is not None
        )
        self.assertGreater(del_para_idx, heading_idx)

    def test_does_not_move_deletion_without_following_plain_body(self):
        """Del-only para after a clean heading is NOT moved when no plain body follows."""
        root = clean_fragment(
            "<w:p>"
            "<w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr>"
            "<w:r><w:t>Section Title</w:t></w:r>"
            "</w:p>"
            "<w:p><w:del w:id=\"1\"><w:r><w:delText>old text</w:delText></w:r></w:del></w:p>"
        )
        paras = root.findall(f".//{W}p")
        # Heading must still be first
        pPr = paras[0].find(f"{W}pPr")
        pStyle = pPr.find(f"{W}pStyle") if pPr is not None else None
        self.assertEqual(pStyle.get(f"{W}val") if pStyle is not None else None, "Heading1")
        self.assertIn("Section Title", "".join(paras[0].itertext()))

    # ── Rule D (inline INS before DEL) with bookmarks in del ─────────────────

    def test_inline_ins_del_suffix_with_bookmarks_in_del(self):
        """Rule D handles del elements containing bookmarks (the actual Section 6 pattern)."""
        root = clean_fragment(
            "<w:p>"
            "<w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr>"
            "<w:ins w:id=\"1\"><w:r><w:t>Representations</w:t></w:r></w:ins>"
            "<w:del w:id=\"2\">"
            "<w:r><w:delText>nature</w:delText></w:r>"
            "<w:bookmarkStart w:id=\"5\" w:name=\"_test\"/>"
            "<w:bookmarkEnd w:id=\"5\"/>"
            "<w:r><w:delText>Representations</w:delText></w:r>"
            "</w:del>"
            "<w:r><w:t xml:space=\"preserve\"> and Warranties</w:t></w:r>"
            "</w:p>"
        )
        p = root.find(f".//{W}p")
        # Insertion must be accepted
        self.assertIsNone(p.find(f".//{W}ins"))
        # "Representations" run removed from del; "nature" remains
        del_texts = [t.text for t in p.findall(f".//{W}delText") if t.text]
        self.assertIn("nature", del_texts)
        self.assertNotIn("Representations", del_texts)
        # Accepted text includes "Representations" and " and Warranties"
        accepted = "".join(p.itertext())
        self.assertIn("Representations", accepted)
        self.assertIn("and Warranties", accepted)

    # ── clean_heading_full_ins_noop: no-fire when no nearby matching del ──────

    def test_heading_full_ins_not_cleaned_when_no_nearby_del_match(self):
        """Heading with only-ins is NOT accepted when no nearby del ends with the same text."""
        root = clean_fragment(
            "<w:p>"
            "<w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr>"
            "<w:ins w:id=\"1\"><w:r><w:t>Section Title</w:t></w:r></w:ins>"
            "</w:p>"
            "<w:p><w:r><w:t>Plain body with no del.</w:t></w:r></w:p>"
        )
        p = root.findall(f".//{W}p")[0]
        # Insertion must remain (no matching del nearby)
        self.assertIsNotNone(p.find(f".//{W}ins"))
        self.assertEqual(
            "".join(p.find(f".//{W}ins").itertext()),
            "Section Title",
        )

    def test_removes_empty_list_paragraph(self):
        """A numbered paragraph with no runs is a Word Compare orphan marker; remove it."""
        root = clean_fragment(
            "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"60\"/></w:numPr></w:pPr></w:p>"
            "<w:p><w:r><w:t>Body.</w:t></w:r></w:p>"
        )
        paras = root.findall(f".//{W}p")
        self.assertEqual(len(paras), 1)
        self.assertEqual("".join(paras[0].itertext()), "Body.")

    def test_removes_whitespace_only_list_paragraph(self):
        """A numbered paragraph whose only run is whitespace is still an empty orphan."""
        root = clean_fragment(
            "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"60\"/></w:numPr></w:pPr>"
            "<w:r><w:t xml:space=\"preserve\"> </w:t></w:r></w:p>"
        )
        self.assertEqual(root.findall(f".//{W}p"), [])

    def test_preserves_list_paragraph_with_bookmark(self):
        """A numbered paragraph anchoring a cross-reference bookmark must be kept."""
        root = clean_fragment(
            "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"3\"/><w:numId w:val=\"45\"/></w:numPr></w:pPr>"
            "<w:bookmarkStart w:id=\"652\" w:name=\"_Ref118098509\"/></w:p>"
        )
        self.assertEqual(len(root.findall(f".//{W}p")), 1)
        self.assertIsNotNone(root.find(f".//{W}bookmarkStart"))

    def test_preserves_list_paragraph_with_deletion(self):
        """A numbered paragraph whose content is a tracked deletion stays visible."""
        root = clean_fragment(
            "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"1\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr>"
            "<w:del w:id=\"85\"><w:r><w:delText>Intentionally omitted. </w:delText></w:r></w:del></w:p>"
        )
        self.assertEqual(len(root.findall(f".//{W}p")), 1)
        self.assertIsNotNone(root.find(f".//{W}del"))

    def test_preserves_list_paragraph_with_text(self):
        """A normal numbered paragraph with visible text is untouched."""
        root = clean_fragment(
            "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr>"
            "<w:r><w:t>Ownership; Assignment.</w:t></w:r></w:p>"
        )
        self.assertEqual(len(root.findall(f".//{W}p")), 1)
        self.assertEqual("".join(root.itertext()).strip(), "Ownership; Assignment.")

    def test_preserves_empty_non_numbered_paragraph(self):
        """An empty paragraph without numbering is an ordinary blank line; keep it."""
        root = clean_fragment("<w:p><w:pPr><w:jc w:val=\"both\"/></w:pPr></w:p>")
        self.assertEqual(len(root.findall(f".//{W}p")), 1)

    def test_preserves_only_paragraph_in_table_cell(self):
        """Never remove a table cell's last paragraph, even if it is an empty list item."""
        root = clean_fragment(
            "<w:tbl><w:tr><w:tc>"
            "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"60\"/></w:numPr></w:pPr></w:p>"
            "</w:tc></w:tr></w:tbl>"
        )
        self.assertEqual(len(root.findall(f".//{W}p")), 1)

    def test_trailing_heading_ins_del_noop_collapses_unchanged_heading(self):
        """Heading text inserted at the tail of a body para and deleted at its heading
        (a few paras later) is an unchanged heading; collapse it to plain text."""
        root = clean_fragment(
            "<w:p>"
            "<w:ins w:id=\"13\">"
            "<w:r><w:t xml:space=\"preserve\">The fees apply. </w:t></w:r>"
            "<w:r><w:t>Reimbursable Expenses</w:t></w:r>"
            "</w:ins>"
            "</w:p>"
            "<w:p><w:del w:id=\"1\"><w:r><w:delText>Old body language.</w:delText></w:r></w:del></w:p>"
            "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr>"
            "<w:bookmarkStart w:id=\"17\" w:name=\"_heading\"/><w:bookmarkEnd w:id=\"17\"/>"
            "<w:del w:id=\"18\"><w:r><w:delText>Reimbursable Expenses</w:delText></w:r></w:del></w:p>"
        )
        paras = root.findall(f".//{W}p")
        # Body insertion kept, but trailing heading run removed
        body_ins_text = "".join(paras[0].find(f".//{W}ins").itertext())
        self.assertIn("The fees apply.", body_ins_text)
        self.assertNotIn("Reimbursable Expenses", body_ins_text)
        # Heading paragraph: no ins/del on the text, shows unchanged text, bookmark kept
        heading = paras[2]
        self.assertEqual(heading.findall(f".//{W}del"), [])
        self.assertEqual(heading.findall(f".//{W}ins"), [])
        self.assertEqual("".join(heading.itertext()), "Reimbursable Expenses")
        self.assertIsNotNone(heading.find(f".//{W}bookmarkStart"))

    def test_trailing_heading_noop_not_applied_without_matching_del(self):
        """No collapse when no nearby heading del matches the trailing insertion."""
        root = clean_fragment(
            "<w:p><w:ins w:id=\"1\"><w:r><w:t>Body. </w:t></w:r><w:r><w:t>Reimbursable Expenses</w:t></w:r></w:ins></w:p>"
            "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr>"
            "<w:del w:id=\"2\"><w:r><w:delText>Different Heading</w:delText></w:r></w:del></w:p>"
        )
        p = root.findall(f".//{W}p")[0]
        self.assertIn("Reimbursable Expenses", "".join(p.find(f".//{W}ins").itertext()))

    def test_trailing_heading_noop_requires_full_del_match(self):
        """A partial (suffix) del match is not a heading no-op; leave it untouched."""
        root = clean_fragment(
            "<w:p><w:ins w:id=\"1\"><w:r><w:t>Body. </w:t></w:r><w:r><w:t>Reimbursable Expenses</w:t></w:r></w:ins></w:p>"
            "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr>"
            "<w:del w:id=\"2\"><w:r><w:delText>Extra Reimbursable Expenses</w:delText></w:r></w:del></w:p>"
        )
        p = root.findall(f".//{W}p")[0]
        self.assertIn("Reimbursable Expenses", "".join(p.find(f".//{W}ins").itertext()))
        heading = root.findall(f".//{W}p")[1]
        self.assertIsNotNone(heading.find(f".//{W}del"))

    def test_trailing_heading_noop_requires_heading_target(self):
        """A matching del in a non-heading paragraph is not collapsed."""
        root = clean_fragment(
            "<w:p><w:ins w:id=\"1\"><w:r><w:t>Body. </w:t></w:r><w:r><w:t>Reimbursable Expenses</w:t></w:r></w:ins></w:p>"
            "<w:p><w:del w:id=\"2\"><w:r><w:delText>Reimbursable Expenses</w:delText></w:r></w:del></w:p>"
        )
        p = root.findall(f".//{W}p")[0]
        self.assertIn("Reimbursable Expenses", "".join(p.find(f".//{W}ins").itertext()))

    def test_trailing_heading_noop_stops_at_real_content(self):
        """A normal paragraph between the insertion and the heading stops the scan."""
        root = clean_fragment(
            "<w:p><w:ins w:id=\"1\"><w:r><w:t>Body. </w:t></w:r><w:r><w:t>Reimbursable Expenses</w:t></w:r></w:ins></w:p>"
            "<w:p><w:r><w:t>Unchanged intervening sentence.</w:t></w:r></w:p>"
            "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr>"
            "<w:del w:id=\"2\"><w:r><w:delText>Reimbursable Expenses</w:delText></w:r></w:del></w:p>"
        )
        p = root.findall(f".//{W}p")[0]
        self.assertIn("Reimbursable Expenses", "".join(p.find(f".//{W}ins").itertext()))

    def test_removes_cosmetic_mid_document_section_break(self):
        """A mid-doc section break with the same page size as the body is a Word Compare
        injection; strip its sectPr but keep the paragraph and the body section."""
        root = clean_fragment(
            "<w:p><w:pPr><w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:pPr>"
            "<w:r><w:t>End of section one.</w:t></w:r></w:p>"
            "<w:p><w:r><w:t>Body.</w:t></w:r></w:p>"
            "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>"
        )
        body = root.find(f"{W}body")
        para_sects = [p for p in body.findall(f"{W}p") if p.find(f"{W}pPr/{W}sectPr") is not None]
        self.assertEqual(para_sects, [])
        # Paragraph text preserved; body-level sectPr preserved
        self.assertEqual("".join(body.findall(f"{W}p")[0].itertext()), "End of section one.")
        self.assertIsNotNone(body.find(f"{W}sectPr"))

    def test_cosmetic_section_break_removal_preserves_bookmark(self):
        """The body's only paragraph is never deleted, so its bookmark stays put."""
        root = clean_fragment(
            "<w:p><w:pPr><w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:pPr>"
            "<w:bookmarkStart w:id=\"5\" w:name=\"_Ref1\"/><w:bookmarkEnd w:id=\"5\"/></w:p>"
            "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>"
        )
        body = root.find(f"{W}body")
        self.assertIsNone(body.find(f"{W}p/{W}pPr/{W}sectPr"))
        self.assertIsNotNone(body.find(f".//{W}bookmarkStart"))
        self.assertEqual(len(body.findall(f"{W}p")), 1)

    def test_removes_empty_section_break_host_paragraph(self):
        """The blank paragraph that carried a cosmetic break goes with it (issue: a stray
        blank line above the next heading)."""
        root = clean_fragment(
            "<w:p><w:r><w:t>Body one.</w:t></w:r></w:p>"
            "<w:p><w:pPr><w:pStyle w:val=\"BodyText\"/>"
            "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:pPr></w:p>"
            "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr>"
            "<w:r><w:t>Definitions</w:t></w:r></w:p>"
            "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>"
        )
        body = root.find(f"{W}body")
        paras = body.findall(f"{W}p")
        self.assertEqual(len(paras), 2)
        self.assertEqual("".join(paras[1].itertext()), "Definitions")

    def test_section_break_host_removal_relocates_bookmark(self):
        """A break host that only anchors a bookmark is removed, the mark moving forward."""
        root = clean_fragment(
            "<w:p><w:r><w:t>Body one.</w:t></w:r></w:p>"
            "<w:p><w:pPr><w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:pPr>"
            "<w:bookmarkStart w:id=\"5\" w:name=\"_Ref1\"/><w:bookmarkEnd w:id=\"5\"/></w:p>"
            "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr>"
            "<w:r><w:t>Definitions</w:t></w:r></w:p>"
            "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>"
        )
        body = root.find(f"{W}body")
        paras = body.findall(f"{W}p")
        self.assertEqual(len(paras), 2)
        heading = paras[1]
        self.assertEqual(heading.find(f"{W}bookmarkStart").get(f"{W}name"), "_Ref1")
        self.assertIsNotNone(heading.find(f"{W}bookmarkEnd"))

    def test_keeps_section_break_host_that_has_text(self):
        """Only a pure spacer is deleted; a break host with text keeps its paragraph."""
        root = clean_fragment(
            "<w:p><w:pPr><w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:pPr>"
            "<w:r><w:t>End of section one.</w:t></w:r></w:p>"
            "<w:p><w:r><w:t>Body.</w:t></w:r></w:p>"
            "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>"
        )
        body = root.find(f"{W}body")
        self.assertEqual(len(body.findall(f"{W}p")), 2)

    def test_keeps_section_break_host_between_two_tables(self):
        """Removing a paragraph between two tables would splice them; keep it."""
        table = (
            "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>{}</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
        )
        root = clean_fragment(
            table.format("first")
            + "<w:p><w:pPr><w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:pPr></w:p>"
            + table.format("second")
            + "<w:p><w:r><w:t>After.</w:t></w:r></w:p>"
            "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>"
        )
        body = root.find(f"{W}body")
        self.assertEqual(len(body.findall(f"{W}tbl")), 2)
        self.assertIsNone(body.find(f"{W}p/{W}pPr/{W}sectPr"))
        # The separator paragraph survives between the two tables.
        self.assertEqual([child.tag for child in body][1], f"{W}p")

    def test_keeps_section_break_host_that_is_cells_last_paragraph(self):
        """A table cell must retain a paragraph, so its only one is never removed."""
        root = clean_fragment(
            "<w:p><w:r><w:t>Before.</w:t></w:r></w:p>"
            "<w:tbl><w:tr><w:tc>"
            "<w:p><w:pPr><w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:pPr></w:p>"
            "</w:tc></w:tr></w:tbl>"
            "<w:p><w:r><w:t>After.</w:t></w:r></w:p>"
            "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>"
        )
        self.assertIsNotNone(root.find(f".//{W}tc/{W}p"))
        self.assertIsNone(root.find(f".//{W}tc/{W}p/{W}pPr/{W}sectPr"))

    def test_preserves_landscape_section_break(self):
        """A section break that changes page orientation/size is genuine; keep it."""
        root = clean_fragment(
            "<w:p><w:pPr><w:sectPr><w:pgSz w:w=\"15840\" w:h=\"12240\" w:orient=\"landscape\"/></w:sectPr></w:pPr>"
            "<w:r><w:t>Landscape exhibit.</w:t></w:r></w:p>"
            "<w:p><w:r><w:t>Body.</w:t></w:r></w:p>"
            "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>"
        )
        self.assertIsNotNone(root.find(f".//{W}p/{W}pPr/{W}sectPr"))

    def test_never_removes_body_level_section(self):
        """The document's final body-level sectPr is always preserved."""
        root = clean_fragment(
            "<w:p><w:r><w:t>Only paragraph.</w:t></w:r></w:p>"
            "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>"
        )
        self.assertIsNotNone(root.find(f"{W}body/{W}sectPr"))

    def test_accepts_image_only_insertion(self):
        """A logo inserted by Word Compare is accepted to a normal, unchanged run."""
        root = clean_fragment(
            "<w:p><w:ins w:id=\"1\"><w:r><w:rPr><w:noProof/></w:rPr>"
            "<w:drawing r:id=\"rId7\"/></w:r></w:ins></w:p>"
        )
        self.assertEqual(root.findall(f".//{W}ins"), [])
        self.assertIsNotNone(root.find(f".//{W}drawing"))

    def test_drops_image_only_deletion(self):
        """A logo deleted by Word Compare is dropped entirely."""
        root = clean_fragment(
            "<w:p><w:del w:id=\"1\"><w:r>"
            "<w:drawing r:id=\"rId8\"/></w:r></w:del></w:p>"
        )
        self.assertEqual(root.findall(f".//{W}del"), [])
        self.assertEqual(root.findall(f".//{W}drawing"), [])

    def test_image_insert_delete_pair_collapses_to_one_unmarked_image(self):
        """The real Phoenix Tower shape: ins-image in one paragraph, del-image in another."""
        root = clean_fragment(
            "<w:p><w:bookmarkStart w:id=\"1\" w:name=\"_x\"/><w:bookmarkEnd w:id=\"1\"/>"
            "<w:ins w:id=\"2\"><w:r><w:drawing r:id=\"rId7\"/></w:r></w:ins></w:p>"
            "<w:p><w:r><w:t>Some heading text.</w:t></w:r></w:p>"
            "<w:p><w:del w:id=\"3\"><w:r><w:drawing r:id=\"rId8\"/></w:r></w:del></w:p>"
        )
        self.assertEqual(root.findall(f".//{W}ins"), [])
        self.assertEqual(root.findall(f".//{W}del"), [])
        self.assertEqual(len(root.findall(f".//{W}drawing")), 1)

    def test_accepts_vml_pict_image_insertion(self):
        """A VML <w:pict> logo insertion is also unwrapped."""
        root = clean_fragment(
            "<w:p><w:ins w:id=\"1\"><w:r><w:pict/></w:r></w:ins></w:p>"
        )
        self.assertEqual(root.findall(f".//{W}ins"), [])
        self.assertIsNotNone(root.find(f".//{W}pict"))

    def test_splits_insertion_mixing_image_and_text(self):
        """An image is never shown as a change, so a mixed revision is split: the image is
        resolved to an unmarked run and the inserted text stays tracked."""
        root = clean_fragment(
            "<w:p><w:ins w:id=\"1\">"
            "<w:r><w:t>New figure: </w:t></w:r><w:r><w:drawing r:embed=\"rId7\"/></w:r>"
            "</w:ins></w:p>"
        )
        ins = root.find(f".//{W}ins")
        self.assertIsNotNone(ins)
        self.assertEqual("".join(ins.itertext()), "New figure: ")
        # The drawing survives, but outside the revision.
        self.assertIsNotNone(root.find(f".//{W}drawing"))
        self.assertIsNone(ins.find(f".//{W}drawing"))

    def test_splits_deletion_mixing_image_and_text(self):
        """The real logo case: Word sweeps a stray character into the picture's <w:del>.
        The image goes; the deleted character stays struck through."""
        root = clean_fragment(
            "<w:p><w:del w:id=\"1\">"
            "<w:r><w:delText>e</w:delText></w:r><w:r><w:drawing r:embed=\"rId8\"/></w:r>"
            "</w:del></w:p>"
        )
        dele = root.find(f".//{W}del")
        self.assertIsNotNone(dele)
        self.assertEqual("".join(dele.itertext()), "e")
        self.assertEqual(root.findall(f".//{W}drawing"), [])

    def test_lifts_image_out_of_a_run_shared_with_text(self):
        """Image and text in one run: the image is lifted into its own unmarked run."""
        root = clean_fragment(
            "<w:p><w:ins w:id=\"1\">"
            "<w:r><w:rPr><w:b/></w:rPr><w:t>Logo</w:t><w:drawing r:embed=\"rId7\"/></w:r>"
            "</w:ins></w:p>"
        )
        ins = root.find(f".//{W}ins")
        self.assertEqual("".join(ins.itertext()), "Logo")
        self.assertIsNone(ins.find(f".//{W}drawing"))
        carrier = root.find(f"{W}body/{W}p/{W}r")
        self.assertIsNotNone(carrier.find(f"{W}drawing"))
        self.assertIsNotNone(carrier.find(f"{W}rPr/{W}b"))  # run properties carried over

    def test_logo_replacement_pair_with_stray_text_collapses(self):
        """End to end: del(stray char + old logo) + ins(new logo) leaves one plain logo."""
        root = clean_fragment(
            "<w:p><w:ins w:id=\"1\"><w:r><w:drawing r:embed=\"rId7\"/></w:r></w:ins></w:p>"
            "<w:p><w:r><w:t>Body.</w:t></w:r></w:p>"
            "<w:p><w:del w:id=\"2\">"
            "<w:r><w:delText>e</w:delText></w:r><w:r><w:drawing r:embed=\"rId8\"/></w:r>"
            "</w:del></w:p>"
        )
        self.assertEqual(root.findall(f".//{W}ins"), [])
        self.assertEqual(len(root.findall(f".//{W}drawing")), 1)
        # The stray deleted character remains visible as a deletion.
        self.assertEqual("".join(root.find(f".//{W}del").itertext()), "e")

    def test_image_revision_with_ole_object_is_untouched(self):
        """OLE objects are not logos; a revision containing one is never split."""
        root = clean_fragment(
            "<w:p><w:del w:id=\"1\">"
            "<w:r><w:delText>Schedule</w:delText></w:r>"
            "<w:r><w:object><o:OLEObject r:id=\"rId5\"/></w:object><w:drawing/></w:r>"
            "</w:del></w:p>"
        )
        self.assertIsNotNone(root.find(f".//{W}del"))
        self.assertIsNotNone(root.find(f".//{O}OLEObject"))
        self.assertIsNotNone(root.find(f".//{W}drawing"))

    def test_image_nested_below_a_run_does_not_spin_the_loop(self):
        """An image the splitter cannot reach must report no change, or clean_root's
        fixed-point loop would never terminate."""
        root = clean_fragment(
            "<w:p><w:ins w:id=\"1\">"
            "<w:hyperlink r:id=\"rId9\"><w:r><w:drawing r:embed=\"rId7\"/></w:r></w:hyperlink>"
            "<w:r><w:t>caption</w:t></w:r>"
            "</w:ins></w:p>"
        )
        # Reaching this assertion at all proves the loop terminated.
        self.assertIsNotNone(root.find(f".//{W}ins"))
        self.assertIsNotNone(root.find(f".//{W}drawing"))

    def test_image_revision_cleaned_in_header_root(self):
        """clean_image_revisions works off a <w:hdr> root (headers are cleaned too)."""
        hdr = ET.fromstring(
            f"<w:hdr xmlns:w=\"{clean_redline.NAMESPACES['w']}\">"
            "<w:p><w:del w:id=\"1\"><w:r><w:drawing/></w:r></w:del></w:p>"
            "</w:hdr>"
        )
        changed = clean_redline.clean_image_revisions(hdr)
        self.assertEqual(changed, 1)
        self.assertEqual(hdr.findall(f".//{W}del"), [])
        self.assertEqual(hdr.findall(f".//{W}drawing"), [])

    def _first_body_p(self, root):
        return root.find(f"{W}body/{W}p")

    def test_body_paragraph_normalized_to_10pt_and_1_15(self):
        """A body run at 12pt / single spacing becomes 10pt / 1.15."""
        root = clean_fragment(
            "<w:p><w:pPr><w:spacing w:line=\"240\" w:lineRule=\"auto\"/></w:pPr>"
            "<w:r><w:rPr><w:sz w:val=\"24\"/><w:szCs w:val=\"24\"/></w:rPr><w:t>Body.</w:t></w:r></w:p>"
        )
        p = self._first_body_p(root)
        spacing = p.find(f"{W}pPr/{W}spacing")
        self.assertEqual(spacing.get(f"{W}line"), "276")
        self.assertEqual(spacing.get(f"{W}lineRule"), "auto")
        sz = p.find(f".//{W}r/{W}rPr/{W}sz")
        self.assertEqual(sz.get(f"{W}val"), "20")
        self.assertEqual(p.find(f".//{W}r/{W}rPr/{W}szCs").get(f"{W}val"), "20")

    def test_inserts_spacing_and_rpr_in_schema_order(self):
        """spacing lands after pStyle and before jc; run gets an rPr with sz=20; order valid."""
        root = clean_fragment(
            "<w:p><w:pPr><w:pStyle w:val=\"BodyText\"/><w:jc w:val=\"both\"/></w:pPr>"
            "<w:r><w:t>Body.</w:t></w:r></w:p>"
        )
        p = self._first_body_p(root)
        order = clean_redline.PPR_ORDER
        names = [c.tag.split("}")[-1] for c in p.find(f"{W}pPr")]
        ranks = [order.index(n) for n in names if n in order]
        self.assertEqual(ranks, sorted(ranks))  # schema order preserved
        self.assertIn("spacing", names)
        self.assertEqual(names.index("pStyle") < names.index("spacing") < names.index("jc"), True)
        run_rpr = p.find(f"{W}r/{W}rPr")
        self.assertIsNotNone(run_rpr)
        self.assertEqual(run_rpr.find(f"{W}sz").get(f"{W}val"), "20")
        # rPr must be the first child of the run
        self.assertEqual(list(p.find(f"{W}r"))[0].tag, f"{W}rPr")

    def test_heading_size_and_spacing_preserved(self):
        """Headings keep their own size and spacing."""
        root = clean_fragment(
            "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr>"
            "<w:r><w:rPr><w:sz w:val=\"28\"/></w:rPr><w:t>Section</w:t></w:r></w:p>"
        )
        p = self._first_body_p(root)
        self.assertEqual(p.find(f".//{W}sz").get(f"{W}val"), "28")
        self.assertIsNone(p.find(f"{W}pPr/{W}spacing"))

    def test_title_preserved(self):
        """The document Title is not normalized."""
        root = clean_fragment(
            "<w:p><w:pPr><w:pStyle w:val=\"Title\"/></w:pPr>"
            "<w:r><w:rPr><w:sz w:val=\"52\"/></w:rPr><w:t>Agreement</w:t></w:r></w:p>"
        )
        p = self._first_body_p(root)
        self.assertEqual(p.find(f".//{W}sz").get(f"{W}val"), "52")

    def test_table_paragraph_normalized(self):
        """Body text inside a table is normalized too."""
        root = clean_fragment(
            "<w:tbl><w:tr><w:tc>"
            "<w:p><w:r><w:rPr><w:sz w:val=\"22\"/></w:rPr><w:t>Cell.</w:t></w:r></w:p>"
            "</w:tc></w:tr></w:tbl>"
        )
        sz = root.find(f".//{W}tbl//{W}r/{W}rPr/{W}sz")
        self.assertEqual(sz.get(f"{W}val"), "20")

    def test_body_formatting_is_idempotent(self):
        """Running clean twice makes no further change on the second pass."""
        body = (
            "<w:p><w:r><w:rPr><w:sz w:val=\"24\"/></w:rPr><w:t>Body.</w:t></w:r></w:p>"
        )
        once, _ = clean_redline.clean_document_xml(
            f"<?xml version=\"1.0\"?><w:document xmlns:w=\"{clean_redline.NAMESPACES['w']}\">"
            f"<w:body>{body}</w:body></w:document>".encode()
        )
        _, changed2 = clean_redline.clean_document_xml(once)
        self.assertEqual(changed2, 0)

    def test_normalize_body_formatting_skips_header_root(self):
        """No <w:body> in a header part → normalize returns 0."""
        hdr = ET.fromstring(
            f"<w:hdr xmlns:w=\"{clean_redline.NAMESPACES['w']}\">"
            "<w:p><w:r><w:rPr><w:sz w:val=\"22\"/></w:rPr><w:t>Header.</w:t></w:r></w:p>"
            "</w:hdr>"
        )
        self.assertEqual(clean_redline.normalize_body_formatting(hdr), 0)
        self.assertEqual(hdr.find(f".//{W}sz").get(f"{W}val"), "22")


    # --- Revision markers in property containers ----------------------------------------

    def test_preserves_tracked_paragraph_mark(self):
        """A deleted paragraph mark carries no text by design, so the empty-wrapper rule
        must not prune it — otherwise a real "this break was removed" edit disappears and
        cannot be rejected."""
        root = clean_fragment(
            "<w:p><w:pPr><w:rPr>"
            "<w:del w:id=\"4\" w:author=\"NU\" w:date=\"2026-01-01T00:00:00Z\"/>"
            "</w:rPr></w:pPr><w:r><w:t>Para one</w:t></w:r></w:p>"
            "<w:p><w:r><w:t>Para two</w:t></w:r></w:p>"
        )
        first = root.find(f"{W}body/{W}p")
        self.assertTrue(clean_redline.paragraph_mark_has_revision(first))

    def test_preserves_inserted_table_row_marker(self):
        """<w:ins> in a <w:trPr> marks an inserted row; it is a marker, not a wrapper."""
        root = clean_fragment(
            "<w:tbl><w:tr><w:trPr><w:ins w:id=\"5\" w:author=\"NU\"/></w:trPr>"
            "<w:tc><w:p><w:r><w:t>New row</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
        )
        self.assertIsNotNone(root.find(f".//{W}trPr/{W}ins"))

    # --- Inserted vs deleted paragraph marks --------------------------------------------
    #
    # Real Word Compare output marks a converted revision's layout noise as *inserted*
    # paragraph marks (12 of them in the Atlas comparison). Deleting such a paragraph is
    # what Reject does to it, so the spacer rules may remove it; a *deleted* mark is the
    # reviewer's to accept and must survive.

    MARK_INS = "<w:pPr><w:rPr><w:ins w:id=\"90\" w:author=\"NU\"/></w:rPr></w:pPr>"
    MARK_DEL = "<w:pPr><w:rPr><w:del w:id=\"91\" w:author=\"NU\"/></w:rPr></w:pPr>"

    def test_paragraph_mark_revision_reports_direction(self):
        ins = ET.fromstring(
            f"<w:p xmlns:w=\"{clean_redline.NAMESPACES['w']}\">{self.MARK_INS}</w:p>"
        )
        dele = ET.fromstring(
            f"<w:p xmlns:w=\"{clean_redline.NAMESPACES['w']}\">{self.MARK_DEL}</w:p>"
        )
        plain = ET.fromstring(f"<w:p xmlns:w=\"{clean_redline.NAMESPACES['w']}\"/>")
        self.assertEqual(clean_redline.paragraph_mark_revision(ins), clean_redline.INS)
        self.assertEqual(clean_redline.paragraph_mark_revision(dele), clean_redline.DEL)
        self.assertIsNone(clean_redline.paragraph_mark_revision(plain))
        # An inserted mark is rejectable; a deleted one is not.
        self.assertTrue(clean_redline.is_rejectable_spacer_paragraph(ins))
        self.assertFalse(clean_redline.is_rejectable_spacer_paragraph(dele))
        # Neither counts as an untracked spacer.
        self.assertFalse(clean_redline.is_spacer_paragraph(ins))
        self.assertFalse(clean_redline.is_spacer_paragraph(dele))

    def test_removes_leading_blank_with_inserted_paragraph_mark(self):
        """The real artifact: the blank above the title arrives as an inserted mark."""
        root = clean_fragment(
            self.LOGO_P
            + f"<w:p>{self.MARK_INS}</w:p>"
            "<w:p><w:pPr><w:pStyle w:val=\"Title\"/></w:pPr>"
            "<w:r><w:t>Agreement</w:t></w:r></w:p>"
        )
        paras = root.find(f"{W}body").findall(f"{W}p")
        self.assertEqual(len(paras), 2)
        self.assertIsNotNone(paras[0].find(f".//{W}drawing"))

    def test_removes_section_break_host_with_inserted_paragraph_mark(self):
        """Same for the spacer that carried a cosmetic section break."""
        root = clean_fragment(
            "<w:p><w:r><w:t>Body one.</w:t></w:r></w:p>"
            "<w:p><w:pPr><w:rPr><w:ins w:id=\"90\" w:author=\"NU\"/></w:rPr>"
            "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:pPr></w:p>"
            "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr>"
            "<w:r><w:t>Definitions</w:t></w:r></w:p>"
            "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>"
        )
        paras = root.find(f"{W}body").findall(f"{W}p")
        self.assertEqual(len(paras), 2)
        self.assertEqual("".join(paras[1].itertext()), "Definitions")

    def test_keeps_section_break_host_with_deleted_paragraph_mark(self):
        """A deleted mark is the reviewer's call — do not pre-empt it."""
        root = clean_fragment(
            "<w:p><w:r><w:t>Body one.</w:t></w:r></w:p>"
            "<w:p><w:pPr><w:rPr><w:del w:id=\"91\" w:author=\"NU\"/></w:rPr>"
            "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr></w:pPr></w:p>"
            "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr>"
            "<w:r><w:t>Definitions</w:t></w:r></w:p>"
            "<w:sectPr><w:pgSz w:w=\"12240\" w:h=\"15840\"/></w:sectPr>"
        )
        body = root.find(f"{W}body")
        self.assertEqual(len(body.findall(f"{W}p")), 3)
        self.assertIsNotNone(body.find(f".//{W}rPr/{W}del"))
        # The cosmetic sectPr still goes; only the paragraph is spared.
        self.assertIsNone(body.find(f"{W}p/{W}pPr/{W}sectPr"))

    def test_removes_stray_table_marked_wholly_inserted(self):
        """The split-off table's cells carry inserted marks; rejecting them drops it."""
        root = clean_fragment(
            self.SIG_TABLE.format(
                "<w:del w:id=\"1\"><w:r><w:delText>Address:  Greenwich, CT</w:delText>"
                "</w:r></w:del>"
            )
            + "<w:p/>"
            "<w:tbl><w:tr><w:tc>"
            f"<w:p>{self.MARK_INS}"
            "<w:ins w:id=\"2\"><w:r><w:t>Address: Greenwich, CT</w:t></w:r></w:ins>"
            "</w:p></w:tc></w:tr></w:tbl>"
        )
        self.assertEqual(len(root.findall(f".//{W}tbl")), 1)
        self.assertEqual(root.findall(f".//{W}ins"), [])

    def test_keeps_stray_table_marked_deleted(self):
        """A table the revision *deletes* stays visible as a deletion."""
        root = clean_fragment(
            self.SIG_TABLE.format(
                "<w:del w:id=\"1\"><w:r><w:delText>Address:  Greenwich, CT</w:delText>"
                "</w:r></w:del>"
            )
            + "<w:p/>"
            "<w:tbl><w:tr><w:tc>"
            f"<w:p>{self.MARK_DEL}"
            "<w:ins w:id=\"2\"><w:r><w:t>Address: Greenwich, CT</w:t></w:r></w:ins>"
            "</w:p></w:tc></w:tr></w:tbl>"
        )
        self.assertEqual(len(root.findall(f".//{W}tbl")), 2)

    def test_merges_split_sentence_across_inserted_mark_spacer(self):
        """The blank between the halves carries an inserted mark in real output."""
        root = clean_fragment(
            "<w:p><w:r><w:t>…the expiration or termination</w:t></w:r></w:p>"
            + f"<w:p>{self.MARK_INS}</w:p>"
            + "<w:p><w:r><w:t>of this Agreement.</w:t></w:r></w:p>"
        )
        paras = root.find(f"{W}body").findall(f"{W}p")
        self.assertEqual(len(paras), 1)
        self.assertEqual(
            "".join(paras[0].itertext()), "…the expiration or termination of this Agreement."
        )

    def test_does_not_merge_when_second_mark_is_deleted(self):
        """A deleted mark on the trailing half is a real edit; leave it visible."""
        root = clean_fragment(
            "<w:p><w:r><w:t>…the expiration or termination</w:t></w:r></w:p>"
            + f"<w:p>{self.MARK_DEL}<w:r><w:t>of this Agreement.</w:t></w:r></w:p>"
        )
        self.assertEqual(len(root.find(f"{W}body").findall(f"{W}p")), 2)

    def test_ignorable_namespace_pruning_converges_in_one_pass(self):
        """Removing a revision orphans the prefix it used, so the mc:Ignorable prune has to
        run after the cleaning loop — otherwise the workflow's second run rewrites the file."""
        xml = (
            "<?xml version=\"1.0\"?>"
            f"<w:document xmlns:w=\"{clean_redline.NAMESPACES['w']}\" "
            f"xmlns:mc=\"{clean_redline.NAMESPACES['mc']}\" "
            f"xmlns:w16du=\"{clean_redline.NAMESPACES['w16du']}\" "
            "mc:Ignorable=\"w16du\"><w:body>"
            "<w:p><w:ins w:id=\"1\" w16du:dateUtc=\"2026-01-01T00:00:00Z\"/></w:p>"
            "<w:p><w:r><w:t>Body.</w:t></w:r></w:p>"
            "</w:body></w:document>"
        ).encode()
        once, _ = clean_redline.clean_document_xml(xml)
        self.assertNotIn(b"w16du", once.split(b"<w:body>")[0])
        _, changed = clean_redline.clean_document_xml(once)
        self.assertEqual(changed, 0)

    # --- Hyperlink revisions (a signature block's e-mail shown as edited) ---------------

    def test_collapses_spacing_only_pair_around_hyperlink_field(self):
        """Word Compare rebuilds a mailto HYPERLINK field, so an unchanged e-mail shows as
        insert+delete differing only in whitespace. Accept the insertion, field and all."""
        root = clean_fragment(
            "<w:p>"
            "<w:ins w:id=\"1\">"
            "<w:r><w:t>Email:</w:t></w:r>"
            "<w:r><w:t xml:space=\"preserve\"> </w:t></w:r>"
            "<w:r><w:fldChar w:fldCharType=\"begin\"/></w:r>"
            "<w:r><w:instrText>HYPERLINK \"mailto:ops@nusearch.com\" \\h</w:instrText></w:r>"
            "<w:r><w:fldChar w:fldCharType=\"separate\"/></w:r>"
            "<w:r><w:t>ops@nusearch.com</w:t></w:r>"
            "<w:r><w:fldChar w:fldCharType=\"end\"/></w:r>"
            "</w:ins>"
            "<w:del w:id=\"2\"><w:r><w:delText>Email:  ops@nusearch.com</w:delText></w:r></w:del>"
            "</w:p>"
        )
        self.assertEqual(root.findall(f".//{W}ins"), [])
        self.assertEqual(root.findall(f".//{W}del"), [])
        # The hyperlink field is preserved verbatim, not flattened to plain text.
        self.assertEqual(len(root.findall(f".//{W}fldChar")), 3)
        self.assertIsNotNone(root.find(f".//{W}instrText"))
        self.assertIn("ops@nusearch.com", "".join(root.itertext()))

    def test_collapses_spacing_only_pair_around_hyperlink_element(self):
        """Same no-op built with <w:hyperlink r:id> instead of a field."""
        root = clean_fragment(
            "<w:p>"
            "<w:ins w:id=\"1\"><w:hyperlink r:id=\"rId9\">"
            "<w:r><w:t>ops@nusearch.com</w:t></w:r></w:hyperlink></w:ins>"
            "<w:del w:id=\"2\"><w:r><w:delText>ops@nusearch.com </w:delText></w:r></w:del>"
            "</w:p>"
        )
        self.assertEqual(root.findall(f".//{W}ins"), [])
        self.assertEqual(root.findall(f".//{W}del"), [])
        self.assertEqual(root.find(f".//{W}hyperlink").get(f"{R}id"), "rId9")

    def test_preserves_pair_when_revision_holds_a_content_control(self):
        """Only hyperlinks are tolerated — an SDT still blocks the pair collapse."""
        root = clean_fragment(
            "<w:p>"
            "<w:ins w:id=\"1\"><w:sdt><w:sdtContent>"
            "<w:r><w:t>Client name</w:t></w:r></w:sdtContent></w:sdt></w:ins>"
            "<w:del w:id=\"2\"><w:r><w:delText>Client name</w:delText></w:r></w:del>"
            "</w:p>"
        )
        self.assertIsNotNone(root.find(f".//{W}ins"))
        self.assertIsNotNone(root.find(f".//{W}del"))

    def test_preserves_pair_when_revision_holds_an_ole_object(self):
        """An embedded OLE object is not hyperlink machinery; leave the revision alone."""
        root = clean_fragment(
            "<w:p>"
            "<w:ins w:id=\"1\"><w:r><w:object><o:OLEObject r:id=\"rId5\"/></w:object>"
            "<w:t>Schedule</w:t></w:r></w:ins>"
            "<w:del w:id=\"2\"><w:r><w:delText>Schedule</w:delText></w:r></w:del>"
            "</w:p>"
        )
        self.assertIsNotNone(root.find(f".//{W}ins"))
        self.assertIsNotNone(root.find(f".//{O}OLEObject"))

    def test_preserves_pair_when_insertion_holds_a_bookmark(self):
        """Bookmarks stay protected: a revision anchoring one is not collapsed."""
        root = clean_fragment(
            "<w:p>"
            "<w:ins w:id=\"1\"><w:bookmarkStart w:id=\"9\" w:name=\"_Ref9\"/>"
            "<w:r><w:t>Fees</w:t></w:r></w:ins>"
            "<w:del w:id=\"2\"><w:r><w:delText>Fees</w:delText></w:r></w:del>"
            "</w:p>"
        )
        self.assertIsNotNone(root.find(f".//{W}ins"))
        self.assertIsNotNone(root.find(f".//{W}bookmarkStart"))

    # --- Leading blank paragraphs above the title ---------------------------------------

    LOGO_P = (
        "<w:p><w:pPr><w:ind w:left=\"3870\"/></w:pPr>"
        "<w:r><w:drawing r:embed=\"rId7\"/></w:r></w:p>"
    )

    def test_removes_leading_blank_paragraphs_above_title(self):
        """Compare keeps both documents' title spacers; fold them into the logo paragraph."""
        root = clean_fragment(
            self.LOGO_P
            + "<w:p><w:pPr><w:pStyle w:val=\"BodyText\"/>"
            "<w:spacing w:before=\"245\"/></w:pPr></w:p>"
            "<w:p><w:pPr><w:pStyle w:val=\"Title\"/><w:spacing w:after=\"320\"/></w:pPr></w:p>"
            "<w:p><w:pPr><w:pStyle w:val=\"Title\"/></w:pPr>"
            "<w:r><w:t>Retained Search Services Agreement</w:t></w:r></w:p>"
        )
        body = root.find(f"{W}body")
        paras = body.findall(f"{W}p")
        self.assertEqual(len(paras), 2)
        self.assertIsNotNone(paras[0].find(f".//{W}drawing"))
        self.assertEqual("".join(paras[1].itertext()), "Retained Search Services Agreement")
        # The largest `after` among the removed spacers moves onto the logo paragraph.
        self.assertEqual(paras[0].find(f"{W}pPr/{W}spacing").get(f"{W}after"), "320")

    def test_leading_blank_removal_defaults_the_title_gap(self):
        """Spacers with no explicit `after` still leave TITLE_GAP_AFTER of space."""
        root = clean_fragment(
            self.LOGO_P
            + "<w:p/>"
            "<w:p><w:pPr><w:pStyle w:val=\"Title\"/></w:pPr>"
            "<w:r><w:t>Agreement</w:t></w:r></w:p>"
        )
        paras = root.find(f"{W}body").findall(f"{W}p")
        self.assertEqual(len(paras), 2)
        self.assertEqual(
            paras[0].find(f"{W}pPr/{W}spacing").get(f"{W}after"),
            clean_redline.TITLE_GAP_AFTER,
        )

    def test_leading_blank_removal_relocates_bookmarks(self):
        """A leading spacer's bookmark moves to the first text paragraph."""
        root = clean_fragment(
            self.LOGO_P
            + "<w:p><w:bookmarkStart w:id=\"3\" w:name=\"_Title\"/>"
            "<w:bookmarkEnd w:id=\"3\"/></w:p>"
            "<w:p><w:r><w:t>Agreement</w:t></w:r></w:p>"
        )
        paras = root.find(f"{W}body").findall(f"{W}p")
        self.assertEqual(len(paras), 2)
        self.assertEqual(paras[1].find(f"{W}bookmarkStart").get(f"{W}name"), "_Title")

    def test_preserves_leading_paragraph_holding_the_logo(self):
        """The logo paragraph has no text but is not a spacer; it bounds the scan."""
        root = clean_fragment(self.LOGO_P + "<w:p><w:r><w:t>Agreement</w:t></w:r></w:p>")
        paras = root.find(f"{W}body").findall(f"{W}p")
        self.assertEqual(len(paras), 2)
        self.assertIsNotNone(paras[0].find(f".//{W}drawing"))

    def test_keeps_leading_blank_with_tracked_paragraph_mark(self):
        """A blank whose paragraph mark is itself a tracked change is a real edit."""
        root = clean_fragment(
            self.LOGO_P
            + "<w:p><w:pPr><w:rPr><w:del w:id=\"4\"/></w:rPr></w:pPr></w:p>"
            "<w:p><w:r><w:t>Agreement</w:t></w:r></w:p>"
        )
        self.assertEqual(len(root.find(f"{W}body").findall(f"{W}p")), 3)

    def test_document_starting_with_text_is_untouched(self):
        """No leading spacers, nothing to do — and no spacing is invented."""
        root = clean_fragment("<w:p><w:r><w:t>Agreement</w:t></w:r></w:p>")
        paras = root.find(f"{W}body").findall(f"{W}p")
        self.assertEqual(len(paras), 1)
        spacing = paras[0].find(f"{W}pPr/{W}spacing")
        self.assertIsNone(spacing.get(f"{W}after"))

    def test_leading_blank_removal_is_idempotent(self):
        """A second clean of the same bytes reports no further change."""
        once, _ = clean_redline.clean_document_xml(
            f"<?xml version=\"1.0\"?><w:document xmlns:w=\"{clean_redline.NAMESPACES['w']}\" "
            f"xmlns:r=\"{clean_redline.NAMESPACES['r']}\"><w:body>"
            f"{self.LOGO_P}<w:p/><w:p/>"
            "<w:p><w:r><w:t>Agreement</w:t></w:r></w:p>"
            "</w:body></w:document>".encode()
        )
        _, changed = clean_redline.clean_document_xml(once)
        self.assertEqual(changed, 0)

    # --- Sentences the conversion split across paragraphs -------------------------------

    def test_merges_sentence_split_across_paragraphs(self):
        """A paragraph stopping mid-sentence rejoins the lower-case one that follows."""
        root = clean_fragment(
            "<w:p><w:r><w:t>…including after the expiration or termination</w:t></w:r></w:p>"
            "<w:p><w:r><w:t>of this Agreement. To be clear, Client pays.</w:t></w:r></w:p>"
        )
        paras = root.find(f"{W}body").findall(f"{W}p")
        self.assertEqual(len(paras), 1)
        self.assertEqual(
            "".join(paras[0].itertext()),
            "…including after the expiration or termination "
            "of this Agreement. To be clear, Client pays.",
        )

    def test_merges_sentence_split_across_two_blank_paragraphs(self):
        """The conversion leaves blanks where the page furniture sat; cross up to two."""
        root = clean_fragment(
            "<w:p><w:r><w:t>…the expiration or termination</w:t></w:r></w:p>"
            "<w:p/><w:p/>"
            "<w:p><w:r><w:t>of this Agreement.</w:t></w:r></w:p>"
        )
        paras = root.find(f"{W}body").findall(f"{W}p")
        self.assertEqual(len(paras), 1)
        self.assertEqual(
            "".join(paras[0].itertext()), "…the expiration or termination of this Agreement."
        )

    def test_does_not_merge_across_three_blank_paragraphs(self):
        """Three blanks is a deliberate gap, not a split sentence."""
        root = clean_fragment(
            "<w:p><w:r><w:t>…the expiration or termination</w:t></w:r></w:p>"
            "<w:p/><w:p/><w:p/>"
            "<w:p><w:r><w:t>of this Agreement.</w:t></w:r></w:p>"
        )
        self.assertEqual(len(root.find(f"{W}body").findall(f"{W}p")), 5)

    def test_does_not_merge_when_next_paragraph_starts_uppercase(self):
        """A new sentence starts a new paragraph; never merge those."""
        root = clean_fragment(
            "<w:p><w:r><w:t>NU will perform the Services in good faith</w:t></w:r></w:p>"
            "<w:p><w:r><w:t>Client will pay all fees when due.</w:t></w:r></w:p>"
        )
        self.assertEqual(len(root.find(f"{W}body").findall(f"{W}p")), 2)

    def test_does_not_merge_when_next_paragraph_starts_with_a_quote(self):
        """Definition paragraphs open with a curly quote, not a lower-case letter."""
        root = clean_fragment(
            "<w:p><w:r><w:t>The following terms apply</w:t></w:r></w:p>"
            "<w:p><w:r><w:t>\u201cCandidate\u201d means any person.</w:t></w:r></w:p>"
        )
        self.assertEqual(len(root.find(f"{W}body").findall(f"{W}p")), 2)

    def test_does_not_merge_when_first_paragraph_ends_in_punctuation(self):
        """A finished sentence was not split."""
        root = clean_fragment(
            "<w:p><w:r><w:t>NU will perform the Services.</w:t></w:r></w:p>"
            "<w:p><w:r><w:t>of this Agreement.</w:t></w:r></w:p>"
        )
        self.assertEqual(len(root.find(f"{W}body").findall(f"{W}p")), 2)

    def test_does_not_merge_a_heading(self):
        """Headings never join running text, even ending without punctuation."""
        root = clean_fragment(
            "<w:p><w:pPr><w:pStyle w:val=\"Heading1\"/></w:pPr>"
            "<w:r><w:t>Fees and Payment Schedule</w:t></w:r></w:p>"
            "<w:p><w:r><w:t>of this Agreement.</w:t></w:r></w:p>"
        )
        self.assertEqual(len(root.find(f"{W}body").findall(f"{W}p")), 2)

    def test_does_not_merge_numbered_paragraphs(self):
        """List items are separate blocks however they end."""
        root = clean_fragment(
            "<w:p><w:pPr><w:numPr><w:numId w:val=\"3\"/></w:numPr></w:pPr>"
            "<w:r><w:t>Checking references as appropriate</w:t></w:r></w:p>"
            "<w:p><w:r><w:t>of this Agreement.</w:t></w:r></w:p>"
        )
        self.assertEqual(len(root.find(f"{W}body").findall(f"{W}p")), 2)

    def test_does_not_merge_paragraphs_inside_a_table(self):
        """Only direct children of <w:body> merge, so cell text is safe."""
        root = clean_fragment(
            "<w:tbl><w:tr><w:tc>"
            "<w:p><w:r><w:t>Payment due on the</w:t></w:r></w:p>"
            "<w:p><w:r><w:t>first business day.</w:t></w:r></w:p>"
            "</w:tc></w:tr></w:tbl>"
        )
        self.assertEqual(len(root.findall(f".//{W}tc/{W}p")), 2)

    def test_does_not_merge_when_paragraph_mark_is_tracked(self):
        """A tracked paragraph mark is a real edit — the reviewer must see it."""
        root = clean_fragment(
            "<w:p><w:pPr><w:rPr><w:del w:id=\"7\"/></w:rPr></w:pPr>"
            "<w:r><w:t>…the expiration or termination</w:t></w:r></w:p>"
            "<w:p><w:r><w:t>of this Agreement.</w:t></w:r></w:p>"
        )
        self.assertEqual(len(root.find(f"{W}body").findall(f"{W}p")), 2)

    # --- No-ops Word Compare spread across a table boundary -----------------------------

    SIG_TABLE = (
        "<w:tbl><w:tr><w:tc><w:p><w:r><w:t>Email: ops@nusearch.com</w:t></w:r></w:p></w:tc>"
        "</w:tr><w:tr><w:tc><w:p>{}</w:p></w:tc></w:tr></w:tbl>"
    )

    def test_collapses_relocated_row_across_tables(self):
        """The revised file split the signature table, so Compare claimed the address moved.
        Keep it where it was and drop the emptied second table."""
        root = clean_fragment(
            self.SIG_TABLE.format(
                "<w:del w:id=\"1\"><w:r><w:delText>Address:  Greenwich, CT</w:delText>"
                "</w:r></w:del>"
            )
            + "<w:p/>"
            "<w:tbl><w:tr><w:tc><w:p>"
            "<w:ins w:id=\"2\"><w:r><w:t>Address: Greenwich, CT</w:t></w:r></w:ins>"
            "</w:p></w:tc></w:tr></w:tbl>"
        )
        tables = root.findall(f".//{W}tbl")
        self.assertEqual(len(tables), 1)
        self.assertEqual(root.findall(f".//{W}ins"), [])
        self.assertEqual(root.findall(f".//{W}del"), [])
        self.assertIn("Address:  Greenwich, CT", "".join(tables[0].itertext()))

    def test_keeps_second_table_holding_other_text(self):
        """Only a table this rule empties is removed."""
        root = clean_fragment(
            self.SIG_TABLE.format(
                "<w:del w:id=\"1\"><w:r><w:delText>Address:  Greenwich, CT</w:delText>"
                "</w:r></w:del>"
            )
            + "<w:p/>"
            "<w:tbl><w:tr>"
            "<w:tc><w:p><w:r><w:t>Keep me.</w:t></w:r></w:p></w:tc>"
            "<w:tc><w:p><w:ins w:id=\"2\"><w:r><w:t>Address: Greenwich, CT</w:t></w:r>"
            "</w:ins></w:p></w:tc>"
            "</w:tr></w:tbl>"
        )
        self.assertEqual(len(root.findall(f".//{W}tbl")), 2)
        self.assertEqual(root.findall(f".//{W}ins"), [])
        self.assertIn("Keep me.", "".join(root.itertext()))

    def test_relocated_table_noop_requires_matching_text(self):
        """A genuinely different insertion in the next table is a real edit."""
        root = clean_fragment(
            self.SIG_TABLE.format(
                "<w:del w:id=\"1\"><w:r><w:delText>Address:  Greenwich, CT</w:delText>"
                "</w:r></w:del>"
            )
            + "<w:p/>"
            "<w:tbl><w:tr><w:tc><w:p>"
            "<w:ins w:id=\"2\"><w:r><w:t>Address: Stamford, CT</w:t></w:r></w:ins>"
            "</w:p></w:tc></w:tr></w:tbl>"
        )
        self.assertIsNotNone(root.find(f".//{W}ins"))
        self.assertIsNotNone(root.find(f".//{W}del"))

    def test_relocated_noop_needs_a_table_boundary(self):
        """Plain adjacent body paragraphs stay the existing cross-paragraph rule's job."""
        root = clean_fragment(
            "<w:p><w:del w:id=\"1\"><w:r><w:delText>Moved text</w:delText></w:r></w:del></w:p>"
            "<w:p><w:ins w:id=\"2\"><w:r><w:t>Moved text</w:t></w:r></w:ins></w:p>"
        )
        # clean_cross_paragraph_noops resolves this one by keeping the insertion.
        self.assertEqual(root.findall(f".//{W}ins"), [])
        self.assertEqual(root.findall(f".//{W}del"), [])
        self.assertIn("Moved text", "".join(root.itertext()))

    # --- Header/footer text flattened into the body -------------------------------------

    @staticmethod
    def _doc_and_footer(body: str, footer: str):
        ns = (
            f"xmlns:w=\"{clean_redline.NAMESPACES['w']}\" "
            f"xmlns:r=\"{clean_redline.NAMESPACES['r']}\""
        )
        doc = ET.fromstring(f"<w:document {ns}><w:body>{body}</w:body></w:document>")
        ftr = ET.fromstring(f"<w:ftr {ns}>{footer}</w:ftr>")
        return doc, ftr

    def test_restores_footer_text_flattened_into_body(self):
        """The revised file had no footer parts, so Compare deleted the footer's text and
        re-inserted it mid-body. Put it back in the footer and drop the body insertion."""
        doc, ftr = self._doc_and_footer(
            "<w:p><w:r><w:t>Body text.</w:t></w:r></w:p>"
            "<w:p><w:ins w:id=\"1\"><w:r><w:t>ID:5155</w:t></w:r></w:ins></w:p>"
            "<w:p><w:r><w:t>More body text.</w:t></w:r></w:p>",
            "<w:p><w:del w:id=\"2\"><w:r><w:delText>ID:5155</w:delText></w:r></w:del></w:p>",
        )
        counts = clean_redline.clean_header_footer_relocations(
            doc, {"word/footer4.xml": ftr}
        )

        self.assertEqual(counts["word/footer4.xml"], 1)
        self.assertEqual(ftr.findall(f".//{W}del"), [])
        self.assertEqual("".join(ftr.itertext()), "ID:5155")
        self.assertEqual(doc.findall(f".//{W}ins"), [])
        self.assertNotIn("ID:5155", "".join(doc.itertext()))
        # The emptied body paragraph goes too, rather than leaving a blank line.
        self.assertEqual(len(doc.findall(f".//{W}p")), 2)

    def test_leaves_footer_deletion_without_a_matching_body_insertion(self):
        """A footer the author really did delete stays deleted."""
        doc, ftr = self._doc_and_footer(
            "<w:p><w:r><w:t>Body text.</w:t></w:r></w:p>",
            "<w:p><w:del w:id=\"2\"><w:r><w:delText>ID:5155</w:delText></w:r></w:del></w:p>",
        )
        counts = clean_redline.clean_header_footer_relocations(
            doc, {"word/footer4.xml": ftr}
        )

        self.assertEqual(counts, {})
        self.assertIsNotNone(ftr.find(f".//{W}del"))

    def test_footer_relocation_requires_matching_text(self):
        """A body insertion of different text is a real edit, not the moved footer."""
        doc, ftr = self._doc_and_footer(
            "<w:p><w:ins w:id=\"1\"><w:r><w:t>ID:9999</w:t></w:r></w:ins></w:p>",
            "<w:p><w:del w:id=\"2\"><w:r><w:delText>ID:5155</w:delText></w:r></w:del></w:p>",
        )
        counts = clean_redline.clean_header_footer_relocations(
            doc, {"word/footer4.xml": ftr}
        )

        self.assertEqual(counts, {})
        self.assertIsNotNone(doc.find(f".//{W}ins"))
        self.assertIsNotNone(ftr.find(f".//{W}del"))


def _document(body: str) -> ET.Element:
    """A <w:document> root for calling a single rule directly, with no other rule running."""
    ns = (
        f"xmlns:w=\"{clean_redline.NAMESPACES['w']}\" "
        f"xmlns:r=\"{clean_redline.NAMESPACES['r']}\""
    )
    return ET.fromstring(f"<w:document {ns}><w:body>{body}</w:body></w:document>")


def _mark(revision: str) -> str:
    """A paragraph mark carrying a tracked change: <w:pPr><w:rPr><w:ins|del/>."""
    return f"<w:pPr><w:rPr><w:{revision} w:id=\"90\"/></w:rPr></w:pPr>"


NUMBERED = "<w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr></w:pPr>"
PHANTOM = (
    "<w:p>" + _mark("del") +
    "<w:ins w:id=\"1\"><w:r><w:t>1.</w:t></w:r><w:r><w:tab/></w:r></w:ins></w:p>"
)
HEADING = f"<w:p>{NUMBERED}<w:r><w:t>Services</w:t></w:r></w:p>"


class PhantomListNumberTests(unittest.TestCase):
    """Compare flattens an auto-number into an inserted literal number in its own paragraph."""

    def test_removes_phantom_number_paragraph(self):
        root = clean_fragment(PHANTOM + HEADING)

        paragraphs = root.findall(f".//{W}p")
        self.assertEqual(len(paragraphs), 1)
        self.assertEqual("".join(paragraphs[0].itertext()), "Services")
        self.assertEqual(root.findall(f".//{W}ins"), [])
        # The heading keeps its automatic numbering.
        self.assertIsNotNone(paragraphs[0].find(f"{W}pPr/{W}numPr"))

    def test_phantom_number_removal_is_idempotent(self):
        once = clean_fragment(PHANTOM + HEADING)
        again = clean_redline.clean_phantom_list_number_paragraphs(once)

        self.assertEqual(again, 0)

    def test_keeps_number_paragraph_whose_mark_is_not_deleted(self):
        """Without the deleted mark it is a real new paragraph, not a merge carrier."""
        body = "<w:p><w:ins w:id=\"1\"><w:r><w:t>1.</w:t></w:r></w:ins></w:p>" + HEADING
        root = _document(body)

        self.assertEqual(clean_redline.clean_phantom_list_number_paragraphs(root), 0)

    def test_keeps_number_paragraph_when_next_is_not_numbered(self):
        """No numPr to merge into means the number is genuinely new text."""
        root = _document(PHANTOM + "<w:p><w:r><w:t>Services</w:t></w:r></w:p>")

        self.assertEqual(clean_redline.clean_phantom_list_number_paragraphs(root), 0)

    def test_keeps_insertion_mixing_a_number_with_words(self):
        """A real edit that happens to start with a number must survive."""
        body = (
            "<w:p>" + _mark("del") +
            "<w:ins w:id=\"1\"><w:r><w:t>1. Extra term</w:t></w:r></w:ins></w:p>"
        ) + HEADING
        root = _document(body)

        self.assertEqual(clean_redline.clean_phantom_list_number_paragraphs(root), 0)

    def test_keeps_number_paragraph_that_is_itself_numbered(self):
        """Its own numPr means it renders a marker: clean_empty_list_paragraphs' business."""
        body = (
            "<w:p><w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr>"
            "<w:rPr><w:del w:id=\"90\"/></w:rPr></w:pPr>"
            "<w:ins w:id=\"1\"><w:r><w:t>1.</w:t></w:r></w:ins></w:p>"
        ) + HEADING
        root = _document(body)

        self.assertEqual(clean_redline.clean_phantom_list_number_paragraphs(root), 0)


class LogoSpacerMarkSwapTests(unittest.TestCase):
    """The letterhead logo's paragraph break is swapped for the one below it."""

    LOGO = "<w:p>" + _mark("ins") + "<w:r><w:drawing/></w:r></w:p>"

    def test_untracks_logo_mark_and_drops_the_emptied_spacer(self):
        root = _document(
            self.LOGO + "<w:p>" + _mark("del") + "</w:p>"
            "<w:p><w:r><w:t>Title</w:t></w:r></w:p>"
        )

        self.assertEqual(clean_redline.clean_logo_spacer_mark_swap(root), 1)
        paragraphs = root.findall(f".//{W}p")
        self.assertEqual(len(paragraphs), 2)
        self.assertIsNone(clean_redline.paragraph_mark_revision(paragraphs[0]))
        self.assertIsNotNone(paragraphs[0].find(f".//{W}drawing"))
        self.assertEqual(root.findall(f".//{W}del"), [])

    def test_logo_swap_is_idempotent(self):
        root = _document(
            self.LOGO + "<w:p>" + _mark("del") + "</w:p>"
            "<w:p><w:r><w:t>Title</w:t></w:r></w:p>"
        )
        clean_redline.clean_logo_spacer_mark_swap(root)

        self.assertEqual(clean_redline.clean_logo_spacer_mark_swap(root), 0)

    def test_keeps_swap_when_the_deleted_paragraph_has_text(self):
        """Real deleted text below the logo is the reviewer's to accept."""
        root = _document(
            self.LOGO +
            "<w:p>" + _mark("del") +
            "<w:del w:id=\"2\"><w:r><w:delText>Old strapline</w:delText></w:r></w:del></w:p>"
        )

        self.assertEqual(clean_redline.clean_logo_spacer_mark_swap(root), 0)

    def test_requires_a_drawing_in_the_inserted_paragraph(self):
        """Two blank paragraphs are clean_inserted_spacer_paragraphs' business, not this."""
        root = _document(
            "<w:p>" + _mark("ins") + "</w:p>" + "<w:p>" + _mark("del") + "</w:p>"
            "<w:p><w:r><w:t>Title</w:t></w:r></w:p>"
        )

        self.assertEqual(clean_redline.clean_logo_spacer_mark_swap(root), 0)


class InsertedSpacerParagraphTests(unittest.TestCase):
    """A converted revision spaces paragraphs with blank lines the original never had."""

    def test_removes_inserted_blank_paragraphs(self):
        root = _document(
            "<w:p><w:r><w:t>First.</w:t></w:r></w:p>"
            "<w:p>" + _mark("ins") + "</w:p>"
            "<w:p>" + _mark("ins") + "</w:p>"
            "<w:p><w:r><w:t>Second.</w:t></w:r></w:p>"
        )

        self.assertEqual(clean_redline.clean_inserted_spacer_paragraphs(root), 2)
        paragraphs = root.findall(f".//{W}p")
        self.assertEqual(len(paragraphs), 2)

    def test_folds_the_removed_gap_onto_the_preceding_paragraph(self):
        """Removing the blank line must not let the page tighten up."""
        root = _document(
            "<w:p><w:r><w:t>First.</w:t></w:r></w:p>"
            "<w:p><w:pPr><w:spacing w:after=\"480\"/><w:rPr><w:ins w:id=\"90\"/></w:rPr>"
            "</w:pPr></w:p>"
            "<w:p><w:r><w:t>Second.</w:t></w:r></w:p>"
        )
        clean_redline.clean_inserted_spacer_paragraphs(root)

        spacing = root.findall(f".//{W}p")[0].find(f"{W}pPr/{W}spacing")
        self.assertEqual(spacing.get(f"{W}after"), "480")

    def test_keeps_a_larger_existing_gap(self):
        root = _document(
            "<w:p><w:pPr><w:spacing w:after=\"960\"/></w:pPr>"
            "<w:r><w:t>First.</w:t></w:r></w:p>"
            "<w:p><w:pPr><w:spacing w:after=\"240\"/><w:rPr><w:ins w:id=\"90\"/></w:rPr>"
            "</w:pPr></w:p>"
            "<w:p><w:r><w:t>Second.</w:t></w:r></w:p>"
        )
        clean_redline.clean_inserted_spacer_paragraphs(root)

        spacing = root.findall(f".//{W}p")[0].find(f"{W}pPr/{W}spacing")
        self.assertEqual(spacing.get(f"{W}after"), "960")

    def test_keeps_untracked_blank_paragraphs(self):
        """An untracked blank line is the reviewer's existing layout, not Compare's noise."""
        root = _document(
            "<w:p><w:r><w:t>First.</w:t></w:r></w:p>"
            "<w:p/>"
            "<w:p><w:r><w:t>Second.</w:t></w:r></w:p>"
        )

        self.assertEqual(clean_redline.clean_inserted_spacer_paragraphs(root), 0)

    def test_keeps_blank_paragraphs_marked_deleted(self):
        """A deleted mark is the reviewer's call to accept — see is_rejectable_spacer_paragraph."""
        root = _document(
            "<w:p><w:r><w:t>First.</w:t></w:r></w:p>"
            "<w:p>" + _mark("del") + "</w:p>"
            "<w:p><w:r><w:t>Second.</w:t></w:r></w:p>"
        )

        self.assertEqual(clean_redline.clean_inserted_spacer_paragraphs(root), 0)

    def test_keeps_a_table_cells_only_paragraph(self):
        """Emptying a cell of paragraphs produces OOXML Word refuses to open."""
        root = _document(
            "<w:p><w:r><w:t>Body.</w:t></w:r></w:p>"
            "<w:tbl><w:tr><w:tc><w:tcPr/>"
            "<w:p>" + _mark("ins") + "</w:p>"
            "</w:tc></w:tr></w:tbl>"
        )

        self.assertEqual(clean_redline.clean_inserted_spacer_paragraphs(root), 0)
        self.assertEqual(len(root.findall(f".//{W}tc/{W}p")), 1)

    def test_inserted_spacer_removal_is_idempotent(self):
        root = _document(
            "<w:p><w:r><w:t>First.</w:t></w:r></w:p>"
            "<w:p>" + _mark("ins") + "</w:p>"
            "<w:p><w:r><w:t>Second.</w:t></w:r></w:p>"
        )
        clean_redline.clean_inserted_spacer_paragraphs(root)

        self.assertEqual(clean_redline.clean_inserted_spacer_paragraphs(root), 0)


class EmptyHeaderFooterMarkTests(unittest.TestCase):
    """A revised file that lost its header parts leaves tracked marks on empty paragraphs."""

    @staticmethod
    def _header(body: str) -> ET.Element:
        ns = f"xmlns:w=\"{clean_redline.NAMESPACES['w']}\""
        return ET.fromstring(f"<w:hdr {ns}>{body}</w:hdr>")

    def test_untracks_an_empty_header_paragraph_mark(self):
        hdr = self._header("<w:p>" + _mark("del") + "</w:p>")

        self.assertEqual(clean_redline.clean_empty_hf_paragraph_marks(hdr), 1)
        self.assertEqual(hdr.findall(f".//{W}del"), [])
        # The paragraph itself stays: a header part needs one.
        self.assertEqual(len(hdr.findall(f".//{W}p")), 1)

    def test_keeps_a_mark_on_a_header_paragraph_with_text(self):
        hdr = self._header(
            "<w:p>" + _mark("del") + "<w:r><w:t>Confidential</w:t></w:r></w:p>"
        )

        self.assertEqual(clean_redline.clean_empty_hf_paragraph_marks(hdr), 0)

    def test_empty_header_mark_removal_is_idempotent(self):
        hdr = self._header("<w:p>" + _mark("ins") + "</w:p>")
        clean_redline.clean_empty_hf_paragraph_marks(hdr)

        self.assertEqual(clean_redline.clean_empty_hf_paragraph_marks(hdr), 0)


class OrphanedMediaTests(unittest.TestCase):
    """Resolving a logo revision orphans the deleted run's relationship and media file."""

    RELS = "http://schemas.openxmlformats.org/package/2006/relationships"
    IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

    def _parts(self, document: str, rels: str, extra=None):
        parts = {
            "word/document.xml": document.encode(),
            "word/_rels/document.xml.rels": (
                f"<Relationships xmlns=\"{self.RELS}\">{rels}</Relationships>"
            ).encode(),
            "word/media/image1.png": b"PNG1",
            "word/media/image2.png": b"PNG2",
        }
        parts.update(extra or {})
        return parts

    def _rel(self, rid: str, target: str) -> str:
        return f"<Relationship Id=\"{rid}\" Type=\"{self.IMAGE}\" Target=\"{target}\"/>"

    DOC_USING_RID7 = (
        "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
        "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
        "<w:body><w:p><w:r><w:drawing><w:blip r:embed=\"rId7\"/></w:drawing></w:r></w:p>"
        "</w:body></w:document>"
    )

    def test_drops_the_orphaned_relationship_and_its_media(self):
        parts = self._parts(
            self.DOC_USING_RID7,
            self._rel("rId7", "media/image1.png") + self._rel("rId8", "media/image2.png"),
        )

        rewritten, dropped, count = clean_redline.prune_orphaned_media(parts)

        self.assertEqual(dropped, {"word/media/image2.png"})
        self.assertEqual(count, 2)  # one relationship + one media file
        self.assertIn("rId7", rewritten["word/_rels/document.xml.rels"].decode())
        self.assertNotIn("rId8", rewritten["word/_rels/document.xml.rels"].decode())

    def test_keeps_media_two_relationships_point_at(self):
        """Only the relationship goes when something else still needs the file."""
        parts = self._parts(
            self.DOC_USING_RID7,
            self._rel("rId7", "media/image1.png") + self._rel("rId8", "media/image1.png"),
        )

        _, dropped, _ = clean_redline.prune_orphaned_media(parts)

        self.assertEqual(dropped, set())

    def test_relationship_ids_are_scoped_to_their_own_part(self):
        """fontTable.xml uses rId8 for a font; that must not protect document.xml's rId8."""
        font_table = (
            "<w:fonts xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" "
            "xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">"
            "<w:font><w:embedRegular r:id=\"rId8\"/></w:font></w:fonts>"
        )
        parts = self._parts(
            self.DOC_USING_RID7,
            self._rel("rId7", "media/image1.png") + self._rel("rId8", "media/image2.png"),
            {"word/fontTable.xml": font_table.encode()},
        )

        _, dropped, _ = clean_redline.prune_orphaned_media(parts)

        self.assertEqual(dropped, {"word/media/image2.png"})

    def test_leaves_everything_alone_when_a_part_will_not_parse(self):
        """Fail closed: a package we cannot read fully is a package we do not touch."""
        parts = self._parts(
            self.DOC_USING_RID7,
            self._rel("rId7", "media/image1.png") + self._rel("rId8", "media/image2.png"),
            {"word/settings.xml": b"<w:settings><broken>"},
        )

        rewritten, dropped, count = clean_redline.prune_orphaned_media(parts)

        self.assertEqual((rewritten, dropped, count), ({}, set(), 0))

    def test_pruning_is_idempotent(self):
        parts = self._parts(
            self.DOC_USING_RID7,
            self._rel("rId7", "media/image1.png") + self._rel("rId8", "media/image2.png"),
        )
        rewritten, dropped, _ = clean_redline.prune_orphaned_media(parts)
        parts.update(rewritten)
        for name in dropped:
            del parts[name]

        self.assertEqual(clean_redline.prune_orphaned_media(parts), ({}, set(), 0))


class OrphanedListNumberDeletionTests(unittest.TestCase):
    """The revised file lost its numbering, so the original's literal number is struck out."""

    IND = "<w:ind w:left=\"720\" w:hanging=\"360\"/>"
    NUMBER_DEL = (
        "<w:del w:id=\"1\"><w:r><w:delText>1.</w:delText></w:r>"
        "<w:r><w:tab/></w:r></w:del>"
    )
    HEADING_RUN = "<w:r><w:t>Services</w:t></w:r>"

    def _heading(self, pPr_extra: str = "", body: str = None) -> ET.Element:
        inner = body if body is not None else (self.NUMBER_DEL + self.HEADING_RUN)
        return _document(
            f"<w:p><w:pPr>{self.IND}{pPr_extra}</w:pPr>{inner}</w:p>"
        )

    def test_restores_the_struck_out_section_number(self):
        root = self._heading()

        self.assertEqual(clean_redline.clean_orphaned_list_number_deletions(root), 1)
        self.assertEqual(root.findall(f".//{W}del"), [])
        self.assertEqual("".join(root.itertext()), "1.Services")
        # The tab that puts the number in the hanging indent survives the conversion.
        self.assertIsNotNone(root.find(f".//{W}tab"))

    def test_restoring_the_number_is_idempotent(self):
        root = self._heading()
        clean_redline.clean_orphaned_list_number_deletions(root)

        self.assertEqual(clean_redline.clean_orphaned_list_number_deletions(root), 0)

    def test_keeps_a_genuine_renumber(self):
        """A real renumber swaps one number for another, so the del is paired with an ins."""
        root = self._heading(
            body="<w:del w:id=\"1\"><w:r><w:delText>1.</w:delText></w:r>"
                 "<w:r><w:tab/></w:r></w:del>"
                 "<w:ins w:id=\"2\"><w:r><w:t>2.</w:t></w:r></w:ins>"
                 + self.HEADING_RUN
        )

        self.assertEqual(clean_redline.clean_orphaned_list_number_deletions(root), 0)

    def test_requires_list_item_indentation(self):
        """Without a hanging indent there is no reserved number slot to argue from."""
        root = _document(
            f"<w:p><w:pPr/>{self.NUMBER_DEL}{self.HEADING_RUN}</w:p>"
        )

        self.assertEqual(clean_redline.clean_orphaned_list_number_deletions(root), 0)

    def test_ignores_a_deleted_figure_inside_a_money_amount(self):
        """"$35,000" -> "$40,000" must stay tracked: no trailing period, not the first child."""
        root = self._heading(
            body="<w:r><w:t>$</w:t></w:r>"
                 "<w:ins w:id=\"2\"><w:r><w:t>40</w:t></w:r></w:ins>"
                 "<w:del w:id=\"1\"><w:r><w:delText>35</w:delText></w:r></w:del>"
                 "<w:r><w:t>,000</w:t></w:r>"
        )

        self.assertEqual(clean_redline.clean_orphaned_list_number_deletions(root), 0)

    def test_requires_a_trailing_period_on_the_number(self):
        root = self._heading(
            body="<w:del w:id=\"1\"><w:r><w:delText>35</w:delText></w:r>"
                 "<w:r><w:tab/></w:r></w:del>" + self.HEADING_RUN
        )

        self.assertEqual(clean_redline.clean_orphaned_list_number_deletions(root), 0)

    def test_requires_a_tab_in_the_deletion(self):
        """Number plus tab is the list layout; a bare number deletion is something else."""
        root = self._heading(
            body="<w:del w:id=\"1\"><w:r><w:delText>1.</w:delText></w:r></w:del>"
                 + self.HEADING_RUN
        )

        self.assertEqual(clean_redline.clean_orphaned_list_number_deletions(root), 0)

    def test_leaves_a_paragraph_that_still_has_its_own_numbering(self):
        """With numPr intact Word supplies the number; the literal one really is redundant."""
        root = self._heading(
            pPr_extra="<w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr>"
        )

        self.assertEqual(clean_redline.clean_orphaned_list_number_deletions(root), 0)

    def test_requires_a_heading_after_the_number(self):
        """A number with nothing after it is not a heading worth keeping it for."""
        root = self._heading(body=self.NUMBER_DEL)

        self.assertEqual(clean_redline.clean_orphaned_list_number_deletions(root), 0)


if __name__ == "__main__":
    unittest.main()
