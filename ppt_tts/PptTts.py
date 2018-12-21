from pathlib import Path, PurePath

from gtts import gTTS
from ppt_tts.exceptions import PptFileDoesNotExist, VoExportDirDoesNotExist

from ppt_tts.PptNotesParser import PptNotesParser


class PptTts:
    """
    Create text-to-speech voice-over files using the notes slides of a PowerPoint file.
    """
    def __init__(self, ppt_file: Path, vo_export_dir: Path):
        if not ppt_file.is_file():
            raise PptFileDoesNotExist()

        if not vo_export_dir.is_dir():
            raise VoExportDirDoesNotExist()

        self.ppt_file = ppt_file
        self.vo_export_dir = vo_export_dir

        notes_parser = PptNotesParser(self.ppt_file)
        self.notes = notes_parser.get_notes()

    def export_vos(self):
        for slide_num, note in enumerate(self.notes):
            save_file = PurePath.joinpath(self.vo_export_dir, f'slide-{slide_num + 1:02d}.mp3')
            tts = gTTS(note)
            tts.save(save_file.as_posix())
