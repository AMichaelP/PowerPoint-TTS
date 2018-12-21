import unittest
import PptTts
from pathlib import Path, PurePath
from PptTts import PptTts

TEST_RESOURCES_DIR = Path(r'C:\Users\User\Desktop')
assert TEST_RESOURCES_DIR.is_dir()

SRC_PPT = PurePath.joinpath(TEST_RESOURCES_DIR, 'test.pptx')
assert SRC_PPT.is_file()

DST_VO_DIR = PurePath.joinpath(TEST_RESOURCES_DIR, 'vo_export')
assert DST_VO_DIR.is_dir()

SLIDE_4_EXPECTED_NOTE_TEXT = 'The definitions shown from RFC 2828 (Internet Security Glossary) ' \
                             'summarize the issues we are discussing. '


class TestPptTts(unittest.TestCase):
    def setUp(self):
        self.ppt_tts = PptTts(SRC_PPT, DST_VO_DIR)

    def test_non_existing_src_raises_AssertionError(self):
        with self.assertRaises(AssertionError):
            ppt_tts = PptTts(Path(''), DST_VO_DIR)

    def test_non_existing_dst_raises_AssertionError(self):
        with self.assertRaises(AssertionError):
            ppt_tts = PptTts(SRC_PPT, Path(r'this\does\not\exist'))

    def test_correct_number_of_notes_slides_extracted(self):
        expected = 6
        actual = len([x for x in self.ppt_tts.notes])
        self.assertEqual(expected, actual)

    def test_get_notes_from_slide_4_return_correct_text(self):
        expected = SLIDE_4_EXPECTED_NOTE_TEXT
        actual = [x for x in self.ppt_tts.notes][3]
        self.assertEqual(expected, actual)

    def test_vos_for_slides_one_three_and_six_are_created_after_vo_export(self):
        self.ppt_tts.export_vos()
        export_files = [x.name for x in DST_VO_DIR.glob('*.mp3')]
        self.assertTrue('slide-01.mp3' in export_files)
        self.assertTrue('slide-03.mp3' in export_files)
        self.assertTrue('slide-06.mp3' in export_files)

