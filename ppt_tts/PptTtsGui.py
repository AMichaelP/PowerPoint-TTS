import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from ppt_tts.PptTts import PptTts
from pathlib import Path
from ppt_tts.exceptions import PptFileDoesNotExist, VoExportDirDoesNotExist, BlankVoExportDir


class PptTtsGui:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title('ppt_tts')
        self.root.resizable(False, False)

        padx = 5
        pady = 5
        label_wrap_length = 100

        self.input_entry_var = tk.StringVar()
        self.export_entry_var = tk.StringVar()

        self.ppt_tts_frame = ttk.Frame(self.root)
        self.ppt_tts_frame.grid(row=0, column=0)

        self.input_label = ttk.Label(self.ppt_tts_frame, text='Input PowerPoint', wraplength=label_wrap_length)
        self.input_label.grid(row=0, column=0, padx=padx, pady=pady)

        self.export_label = ttk.Label(self.ppt_tts_frame, text='Voice-over Export Directory',
                                      wraplength=label_wrap_length)
        self.export_label.grid(row=1, column=0, padx=padx, pady=pady)

        self.input_entry = ttk.Entry(self.ppt_tts_frame, textvariable=self.input_entry_var)
        self.input_entry.grid(row=0, column=1, padx=padx, pady=pady)

        self.export_entry = ttk.Entry(self.ppt_tts_frame, textvariable=self.export_entry_var)
        self.export_entry.grid(row=1, column=1, padx=padx, pady=pady)

        self.input_browse_btn = ttk.Button(self.ppt_tts_frame, text='Browse...', command=self.browse_ppt_file)
        self.input_browse_btn.grid(row=0, column=2, padx=padx, pady=pady)

        self.export_browse_btn = ttk.Button(self.ppt_tts_frame, text='Browse...', command=self.browse_export_dir)
        self.export_browse_btn.grid(row=1, column=2, padx=padx, pady=pady)

        self.start_btn = ttk.Button(self.ppt_tts_frame, text='Start', command=self.export_vos)
        self.start_btn.grid(row=3, column=2, padx=padx, pady=pady)

        self.status_label = ttk.Label(self.ppt_tts_frame, text='Ready')
        self.status_label.grid(row=3, columnspan=2, sticky='SW', padx=padx, pady=pady)

        self.root.mainloop()

    def export_vos(self):
        ppt_file = Path(self.input_entry_var.get())
        vo_export_dir = Path(self.export_entry_var.get())

        try:
            if not vo_export_dir:
                raise BlankVoExportDir
            ppt_tts = PptTts(ppt_file, vo_export_dir)
            self.update_status('Exporting')
            self.root.update()
            ppt_tts.export_vos()
        except PptFileDoesNotExist:
            messagebox.showerror('Invalid PowerPoint Path', 'The given PowerPoint file could not be located.')
        except VoExportDirDoesNotExist:
            messagebox.showerror('Invalid Export Path', 'The given export directory could not be located.')
        except BlankVoExportDir:
            messagebox.showerror('Invalid Export Path', 'The export path can not be blank.')

        self.update_status('Done')

    def update_status(self, text: str):
        self.status_label.configure(text=text)

    def browse_ppt_file(self):
        ppt_file = filedialog.askopenfile(initialdir='/', title='Select a PowerPoint file',
                                          filetypes=(('PowerPoint files', ['*.ppt', '*.pptx']), ('all files', '*.*')))
        self.input_entry_var.set(ppt_file.name)

    def browse_export_dir(self):
        export_dir = filedialog.askdirectory(initialdir='/', title='Select an export directory')
        self.export_entry_var.set(export_dir)
