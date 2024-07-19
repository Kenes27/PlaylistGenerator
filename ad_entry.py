import tkinter as tk
from tkinter import ttk, filedialog
from pydub import AudioSegment
import os

class AdEntry:

    def __init__(self, parent, index, ad_data=None):
        self.parent = parent
        self.index = index
        self.frame = tk.Frame(parent.ad_frame)
        self.frame.pack(padx=5, pady=5)
        self.create_widgets(ad_data)
        parent.ad_frame_list.append(self.frame)

    def create_widgets(self, ad_data):
        tk.Label(self.frame, text=str(self.index + 1) + ". Файл рекламы:").grid(row=0, column=0, padx=5, pady=5)
        self.ad_file_entry = tk.Entry(self.frame, width=30)
        self.ad_file_entry.grid(row=0, column=1, padx=5, pady=5)
        if ad_data:
            self.ad_file_entry.insert(0, ad_data[0])
        self.dur_label = tk.Label(self.frame, text='')
        self.dur_label.grid(row=0, column=5, padx=5, pady=5)
        if ad_data and os.path.isfile(ad_data[0]) and os.path.splitext(ad_data[0])[-1].lower() in {".mp3", ".wav", ".flac", ".m4a", ".ogg"}:
            self.dur_label.config(text='Продолжительность: ' + str(int(AudioSegment.from_file(ad_data[0]).duration_seconds)) + ' сек.')

        tk.Button(self.frame, text="Обзор", command=self.browse_ad_file).grid(row=0, column=2, padx=5, pady=5)

        tk.Label(self.frame, text="Повторы:").grid(row=0, column=3, padx=5, pady=5)
        self.ad_repeat_entry = ttk.Combobox(self.frame, values=["20", "15", "10", "5"])
        self.ad_repeat_entry.set("20")
        if ad_data:
            self.ad_repeat_entry.set(ad_data[1])
        self.ad_repeat_entry.grid(row=0, column=4, padx=5, pady=5)

        self.ad_file_entry.bind("<KeyRelease>", lambda event: self.parent.update_load())
        self.ad_repeat_entry.bind("<<ComboboxSelected>>", lambda event: self.parent.update_load())

        self.add_move_buttons()

    def browse_ad_file(self):
        file = filedialog.askopenfilename(filetypes=[("Audio files", "*.mp3 *.wav *.flac *.m4a *.ogg")])
        if file:
            self.ad_file_entry.delete(0, tk.END)
            self.ad_file_entry.insert(0, file)
            self.dur_label.config(text='Продолжительность: ' + str(int(AudioSegment.from_file(file).duration_seconds)) + ' сек.')
        self.parent.update_load()

    def add_move_buttons(self):
        tk.Button(self.frame, text="Вверх", command=lambda: self.parent.move_advertisement_up(self.index), bg='lightgreen').grid(row=0, column=6, padx=5, pady=5)
        tk.Button(self.frame, text="Вниз", command=lambda: self.parent.move_advertisement_down(self.index), bg='lightblue').grid(row=0, column=7, padx=5, pady=5)

    def update_index(self, new_index):
        self.index = new_index
        self.frame.winfo_children()[0].config(text=f"{self.index + 1}. Файл рекламы:")
