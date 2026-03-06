from __future__ import annotations

from pathlib import Path
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET

from h2hx_py import convert_file


FIXTURES_DIR = Path(__file__).resolve().parent / "fixtures"


class ConversionTest(unittest.TestCase):
    def test_convert_creates_hwpx_package(self) -> None:
        source = FIXTURES_DIR / "arc.hwp"
        with tempfile.TemporaryDirectory() as tmpdir:
            output = Path(tmpdir) / "from.hwpx"
            result = convert_file(source, output)
            self.assertTrue(result.output_path.exists())
            with zipfile.ZipFile(result.output_path) as archive:
                names = set(archive.namelist())
            self.assertIn("mimetype", names)
            self.assertIn("version.xml", names)
            self.assertIn("Contents/content.hpf", names)
            self.assertIn("Contents/header.xml", names)
            self.assertIn("Contents/section0.xml", names)
            self.assertIn("settings.xml", names)

    def test_table_controls_stay_nested_in_cells(self) -> None:
        source = FIXTURES_DIR / "table.hwp"
        with tempfile.TemporaryDirectory() as tmpdir:
            output = Path(tmpdir) / "from.hwpx"
            result = convert_file(source, output)
            with zipfile.ZipFile(result.output_path) as archive:
                section_root = ET.fromstring(archive.read("Contents/section0.xml"))
            ns = {"hp": "http://www.hancom.co.kr/hwpml/2011/paragraph"}
            self.assertGreater(len(section_root.findall(".//hp:tbl", ns)), 0)
            self.assertGreater(len(section_root.findall(".//hp:tbl//hp:p", ns)), 0)
            self.assertLess(
                len(section_root.findall("./hp:p", ns)),
                len(section_root.findall(".//hp:p", ns)),
            )


if __name__ == "__main__":
    unittest.main()
