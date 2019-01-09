# PowerPoint TTS
Create text-to-speech voice-over files using the slide notes of a PowerPoint file. Can be used from the terminal, 
or by using the GUI.

![ppt_tts GUI](https://github.com/AMichaelP/ppt_tts/blob/master/docs/images/gui_screenshot.png)

## Setup

Install [gTTS](https://pypi.org/project/gTTS/) and clone the ppt_tts repository, or download it as a ZIP:

`pip install gtts`

`git clone https://github.com/AMichaelP/ppt_tts.git`

## Usage

Running ppt_tts from the terminal:

`python ppt_tts.py C:\Users\User\Documents\my_presentation.pptx C:\Users\User\Desktop\vo_export`

Starting ppt_tts in GUI mode:

`python ppt_tts.py --gui`
