import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Iterable, List
from zipfile import ZipFile


class PptNotesParser:
    """
    Extract slide notes from a PowerPoint file.
    """
    def __init__(self, ppt_file: Path):
        self.ppt_file = ppt_file

    def get_notes(self) -> Iterable[str]:
        notes_xml = self.get_notes_xml_from_ppt(self.ppt_file)
        for x in notes_xml:
            xml = ET.fromstring(x)
            slide_text = self.get_slide_text_from_xml(xml)
            yield slide_text

    def get_notes_xml_from_ppt(self, ppt_file: Path) -> Iterable[str]:
        ppt_zip = ZipFile(ppt_file)
        note_files = self.get_note_slides_file_name_from_zip_file(ppt_zip)
        for note in note_files:
            yield ppt_zip.read(note).decode()

    @staticmethod
    def get_slide_text_from_xml(xml: ET.Element) -> str:
        text = [x.text for x in xml.iter() if x.text][1]
        return text

    @staticmethod
    def get_note_slides_file_name_from_zip_file(zip_file: ZipFile) -> List[str]:
        note_files = [x for x in zip_file.namelist() if x.startswith('ppt/notesSlides/notesSlide')]
        return note_files
