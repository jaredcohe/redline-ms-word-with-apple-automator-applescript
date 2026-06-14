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

    def test_preserves_empty_revision_with_drawing_relationship(self):
        root = clean_fragment(
            "<w:p><w:ins w:id=\"1\"><w:r><w:drawing r:id=\"rId5\"/></w:r></w:ins></w:p>"
        )

        self.assertEqual(len(root.findall(f".//{W}ins")), 1)
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

        paras = root.findall(f".//{W}p")
        heading_p = paras[2]

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


if __name__ == "__main__":
    unittest.main()
