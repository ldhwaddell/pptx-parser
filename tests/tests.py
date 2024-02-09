import io
import unittest

from pptx_parser.pptx_parser import PptxParser

from contextlib import redirect_stdout


class Test(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        parser_pptx = PptxParser(dir="tests/powerpoints/pptx", recursive=True)
        parser_image = PptxParser(dir="tests/powerpoints/image", recursive=True)
        parser_audio = PptxParser(dir="tests/powerpoints/audio", recursive=True)

        # Redirect output to suppress print
        with io.StringIO() as buf, redirect_stdout(buf):
            cls.parser_pptx_results = parser_pptx.parse()
            cls.parser_image_results = parser_image.parse()
            cls.parser_audio_results = parser_audio.parse()

    def test_parse_pptx(self):
        # Ensure 1 ppt was found
        self.assertEqual(len(self.parser_pptx_results), 1)

    def test_parse_pptx_slides(self):
        # Ensure both slides and title were found
        self.assertEqual(len(self.parser_pptx_results[0].keys()), 3)

    def test_parse_pptx_text(self):
        # Ensure text properly extracted
        slide1 = self.parser_pptx_results[0]["slide1"]
        slide2 = self.parser_pptx_results[0]["slide2"]

        self.assertEqual(slide1["slide_text"], "Title Subtitle")
        self.assertEqual(slide2["slide_text"], "Slide Title 1 2")

    # def test_parse_pptx_images(self):
    #     # Ensure images are properly found
    #     print(self.parser_image_results)

    def test_parse_pptx_audio(self):
        # Ensure audio is properly transcribed
        slide2 = self.parser_audio_results[0]["slide2"]

        self.assertEqual(slide2["transcription"], " This is a test.")


# Run the tests
if __name__ == "__main__":
    unittest.main()
