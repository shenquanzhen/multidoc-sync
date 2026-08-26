import tempfile
import unittest
from pathlib import Path

from multi_document_viewer import is_supported, parse_initial_paths
from native_office import document_family


class CoreTests(unittest.TestCase):
    def test_supported_extensions(self):
        for name in ["a.pdf", "a.doc", "a.docx", "a.ppt", "a.pptx"]:
            self.assertTrue(is_supported(name))
        self.assertFalse(is_supported("a.xlsx"))

    def test_document_family_accepts_same_office_family(self):
        self.assertEqual(document_family(["a.doc", "b.docx", "c.docx"]), "word")
        self.assertEqual(document_family(["a.ppt", "b.pptx", "c.pptx"]), "powerpoint")
        self.assertEqual(document_family(["a.pdf", "b.pdf", "c.pdf"]), "pdf")

    def test_document_family_rejects_mixed_types(self):
        self.assertIsNone(document_family(["a.pdf", "b.docx", "c.pptx"]))

    def test_parse_paths_preserves_unicode_and_spaces(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "中文 文件.pdf"
            parsed = parse_initial_paths([str(path)])
            self.assertEqual(parsed, [str(path)])


if __name__ == "__main__":
    unittest.main()
