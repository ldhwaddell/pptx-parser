import io
import unittest

from pptx_parser.pptx_parser import PptxParser

from contextlib import redirect_stdout


class Test(unittest.TestCase):

    def test_get_pptx_files(self):
        parser = PptxParser(dir="tests/powerpoints", recursive=False)
        # Redirect stdout to suppress print output
        with io.StringIO() as buf, redirect_stdout(buf):
            result = parser._PptxParser__get_pptx_files()
            self.assertEqual(len(result), 2)

    def test_get_pptx_files_recursive(self):
        parser = PptxParser(dir="tests/powerpoints", recursive=True)

        with io.StringIO() as buf, redirect_stdout(buf):
            result = parser._PptxParser__get_pptx_files()
            self.assertEqual(len(result), 3)

    def test_get_pptx_files_bad_path(self):
        parser = PptxParser(dir="nonexistent_directory")

        with self.assertRaises(NotADirectoryError):
            parser._PptxParser__get_pptx_files()

    def test_get_pptx_files_empty_path(self):
        parser = PptxParser(dir="tests/powerpoints/empty_dir")

        with self.assertRaises(FileNotFoundError):
            parser._PptxParser__get_pptx_files()


# Run the tests
if __name__ == "__main__":
    unittest.main()
