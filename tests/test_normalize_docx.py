import unittest
import xml.etree.ElementTree as ET

import normalize_docx


W = normalize_docx.W


def normalize_fragment(body: str) -> ET.Element:
    xml = f"""<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="{normalize_docx.W_NS}">
  <w:body>{body}</w:body>
</w:document>"""
    normalized, _ = normalize_docx.normalize_xml(xml.encode("utf-8"))
    return ET.fromstring(normalized)


class NormalizeDocxTests(unittest.TestCase):
    def test_unwraps_inline_content_control(self):
        root = normalize_fragment(
            "<w:p><w:r><w:t>between </w:t></w:r>"
            "<w:sdt><w:sdtPr><w:alias w:val=\"Company\"/></w:sdtPr>"
            "<w:sdtContent><w:r><w:t>Company Name</w:t></w:r></w:sdtContent>"
            "</w:sdt><w:r><w:t> and Vendor</w:t></w:r></w:p>"
        )

        self.assertEqual(root.findall(f".//{W}sdt"), [])
        self.assertEqual("".join(root.itertext()).strip(), "between Company Name and Vendor")

    def test_unwraps_block_content_control(self):
        root = normalize_fragment(
            "<w:sdt><w:sdtPr/><w:sdtContent>"
            "<w:p><w:r><w:t>Visible paragraph</w:t></w:r></w:p>"
            "</w:sdtContent></w:sdt>"
        )

        self.assertEqual(root.findall(f".//{W}sdt"), [])
        self.assertEqual("".join(root.itertext()).strip(), "Visible paragraph")


if __name__ == "__main__":
    unittest.main()
