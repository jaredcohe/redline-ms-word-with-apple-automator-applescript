import os
import tempfile
import unittest

import next_version_name as nv


class BumpTests(unittest.TestCase):
    def test_appends_minor_when_none(self):
        self.assertEqual(nv.bump("v7"), "v7.1")

    def test_bumps_existing_minor(self):
        self.assertEqual(nv.bump("v7.1"), "v7.2")

    def test_underscore_separator(self):
        # The real bug: v7_260718_... must increment (trailing \b fails on '_').
        self.assertEqual(
            nv.bump("v7_260718_Raven_Rosey_CDA_-_Fully_Mutual"),
            "v7.1_260718_Raven_Rosey_CDA_-_Fully_Mutual",
        )

    def test_space_separator_still_works(self):
        self.assertEqual(nv.bump("v1 260717_Raven Rosey CDA"), "v1.1 260717_Raven Rosey CDA")

    def test_decimal_underscore(self):
        self.assertEqual(nv.bump("v7.1_260718_Fully_Mutual"), "v7.2_260718_Fully_Mutual")

    def test_capital_v_and_multidigit(self):
        self.assertEqual(nv.bump("V10_Big"), "V10.1_Big")

    def test_no_version_token_returns_none(self):
        self.assertIsNone(nv.bump("Raven Rosey CDA"))


class NextFreeNameTests(unittest.TestCase):
    def _touch(self, directory, base):
        open(os.path.join(directory, base + nv.EXT), "w").close()

    def test_single_bump_when_no_collision(self):
        with tempfile.TemporaryDirectory() as d:
            self.assertEqual(
                nv.next_free_name(d, "v7_260718_Fully_Mutual"),
                "v7.1_260718_Fully_Mutual",
            )

    def test_bumps_past_existing_output(self):
        with tempfile.TemporaryDirectory() as d:
            self._touch(d, "v7.1_260718_Fully_Mutual")
            self.assertEqual(
                nv.next_free_name(d, "v7_260718_Fully_Mutual"),
                "v7.2_260718_Fully_Mutual",
            )

    def test_bumps_past_multiple_existing_outputs(self):
        with tempfile.TemporaryDirectory() as d:
            self._touch(d, "v7.1_x")
            self._touch(d, "v7.2_x")
            self.assertEqual(nv.next_free_name(d, "v7_x"), "v7.3_x")

    def test_no_version_falls_back_to_redline_suffix(self):
        with tempfile.TemporaryDirectory() as d:
            self.assertEqual(nv.next_free_name(d, "Raven Rosey CDA"), "Raven Rosey CDA redline")

    def test_no_version_suffix_counter_on_collision(self):
        with tempfile.TemporaryDirectory() as d:
            self._touch(d, "Raven Rosey CDA redline")
            self.assertEqual(
                nv.next_free_name(d, "Raven Rosey CDA"),
                "Raven Rosey CDA redline 2",
            )


if __name__ == "__main__":
    unittest.main()
