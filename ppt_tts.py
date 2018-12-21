import argparse
from pathlib import Path

from ppt_tts.PptTts import PptTts


def main():
    parser = argparse.ArgumentParser(description='Create text-to-speech voice-over files using the slide notes of a '
                                                 'PowerPoint file')
    parser.add_argument('ppt_file', type=Path, help='Input PowerPoint file')
    parser.add_argument('vo_export_dir', type=Path, help='Export directory for voice-over files')

    args = parser.parse_args()

    ppt_file: Path = args.ppt_file
    vo_export_dir: Path = args.vo_export_dir

    ppt_tts = PptTts(ppt_file, vo_export_dir)

    print(f'Exporting voice-over files to {vo_export_dir.absolute()}')
    ppt_tts.export_vos()
    print('Done.')


if __name__ == '__main__':
    main()
