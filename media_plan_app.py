import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel, Label
import json
from ad_entry import AdEntry
from media_plan_generator import MediaPlanGenerator
from utils import hour_to_seconds
from pydub import AudioSegment


class MediaPlanApp:

    def __init__(self, root):
        self.waiting_window = None
        self.root = root
        self.root.title("Генератор медиаплана")

        self.ad_dur = []
        self.ad_repeat = []
        self.ad_repeat_combo = []
        self.ad_name = []
        self.ad_files = []
        self.music_files = []
        self.ad_frame_list = []
        self.number = 0

        self.setup_initial_gui()

    def setup_initial_gui(self):

        exe_path = os.path.abspath(sys.argv[0])
        exe_dir = os.path.dirname(exe_path)

        json_data = None

        if os.path.isfile(exe_dir + '/data.json'):
            with open(exe_dir + '/data.json', encoding='ascii') as f:
                json_data = json.load(f)

        self.frame = tk.Frame(self.root, padx=10, pady=10)
        self.frame.pack(padx=10, pady=10)

        tk.Label(self.frame, text="Папка с музыкальными файлами:", anchor='w').grid(row=0, column=0, sticky=tk.W,
                                                                                    padx=5, pady=5)
        self.music_files_entry = tk.Entry(self.frame, width=50)
        self.music_files_entry.grid(row=0, column=1, padx=5, pady=5)
        if json_data != None:
            self.music_files_entry.insert(0, json_data["music"])
        tk.Button(self.frame, text="Обзор", command=self.browse_music_files).grid(row=0, column=2, padx=5, pady=5)

        tk.Label(self.frame, text="Начальное время:", anchor='w').grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.start_time_entry = tk.Entry(self.frame, width=10)
        if json_data != None and json_data["start"] != "":
            self.start_time_entry.insert(0, json_data["start"])
        else:
            self.start_time_entry.insert(0, "09:00:00")
        self.start_time_entry.grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.start_time_entry.bind("<KeyRelease>", lambda event: self.update_load())

        tk.Label(self.frame, text="Конечное время:", anchor='w').grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)
        self.end_time_entry = tk.Entry(self.frame, width=10)
        if json_data != None and json_data["end"] != "":
            self.end_time_entry.insert(0, json_data["end"])
        else:
            self.end_time_entry.insert(0, "16:00:00")
        self.end_time_entry.grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)
        self.end_time_entry.bind("<KeyRelease>", lambda event: self.update_load())

        self.load = tk.Label(self.frame, text="", anchor='w')
        self.load.grid(row=1, column=2, sticky=tk.W, padx=5, pady=5)

        tk.Button(self.frame, text="Добавить рекламу", command=self.add_advertisement).grid(row=3, column=0, pady=10)

        self.canvas = tk.Canvas(self.root, borderwidth=0, width=500, height=300)
        self.ad_frame = tk.Frame(self.canvas, padx=10, pady=10)
        self.scrollbar = tk.Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.create_window((4, 4), window=self.ad_frame, anchor="nw", tags="self.ad_frame")

        self.ad_frame.bind("<Configure>", lambda event, canvas=self.canvas: self.on_frame_configure(canvas))

        tk.Button(self.frame, text="Удалить рекламу", command=self.delete_advertisement).grid(row=3, column=1, pady=10)
        tk.Button(self.frame, text="Сгенерировать медиаплан", command=self.generate_media_plan).grid(row=3, column=2, pady=10)

        if json_data != None and "ad" in json_data.keys():
            for x in json_data["ad"]:
                self.add_advertisement_init(x[0], x[1])

    def on_frame_configure(self, canvas):
        canvas.configure(scrollregion=canvas.bbox("all"))

    def browse_music_files(self):
        files = filedialog.askdirectory()
        if files:
            self.music_files_entry.delete(0, tk.END)
            self.music_files_entry.insert(0, files)

    def move_advertisement_up(self, index):
        if index > 0:
            self.swap_advertisements(index, index - 1)

    def move_advertisement_down(self, index):
        if index < len(self.ad_frame_list) - 1:
            self.swap_advertisements(index, index + 1)

    def swap_advertisements(self, index1, index2):
        self.ad_frame_list[index1], self.ad_frame_list[index2] = self.ad_frame_list[index2], self.ad_frame_list[index1]
        self.ad_files[index1], self.ad_files[index2] = self.ad_files[index2], self.ad_files[index1]
        self.ad_repeat_combo[index1], self.ad_repeat_combo[index2] = self.ad_repeat_combo[index2], self.ad_repeat_combo[index1]

        for widget in self.ad_frame.winfo_children():
            widget.pack_forget()
        for i, frame in enumerate(self.ad_frame_list):
            frame.pack(padx=5, pady=5)
            self.update_move_buttons(i)

        for i, frame in enumerate(self.ad_frame_list):
            frame.winfo_children()[0].config(text=f"{i + 1}. Файл рекламы:")

        self.update_load()

    def update_move_buttons(self, index):
        frame = self.ad_frame_list[index]
        move_up_button = frame.winfo_children()[6]
        move_down_button = frame.winfo_children()[7]

        move_up_button.config(command=lambda idx=index: self.move_advertisement_up(idx))
        move_down_button.config(command=lambda idx=index: self.move_advertisement_down(idx))

    def add_advertisement(self):
        self.number += 1
        ad_entry = AdEntry(self, self.number - 1)
        self.ad_files.append(ad_entry.ad_file_entry)
        self.ad_repeat_combo.append(ad_entry.ad_repeat_entry)
        self.update_load()

    def add_advertisement_init(self, path, repeat):
        self.number += 1
        ad_entry = AdEntry(self, self.number - 1, ad_data=[path, repeat])
        self.ad_files.append(ad_entry.ad_file_entry)
        self.ad_repeat_combo.append(ad_entry.ad_repeat_entry)
        self.update_load()

    def delete_advertisement(self):
        if len(self.ad_frame_list) == 0:
            messagebox.showerror("Ошибка", "Нет рекламы, которую можно удалить.")
        else:
            self.number -= 1
            for widget in self.ad_frame_list[-1].winfo_children():
                widget.destroy()
            self.ad_frame_list[-1].destroy()
            self.ad_frame_list.pop()
            self.ad_repeat_combo.pop()
            self.ad_files.pop()
            self.update_load()

    def show_waiting_message(self):
        self.waiting_window = Toplevel(self.root)
        self.waiting_window.title("Пожалуйста, подождите")
        self.waiting_window.geometry("300x100")
        Label(self.waiting_window, text="Медиаплан генерируется, пожалуйста, подождите...").pack(expand=True)
        self.waiting_window.grab_set()
        self.root.update_idletasks()

    def close_waiting_message(self):
        if hasattr(self, 'waiting_window') and self.waiting_window:
            self.waiting_window.destroy()

    def generate_media_plan(self):
        self.show_waiting_message()
        self.root.after(100, self._generate_media_plan)

    def _generate_media_plan(self):
        generator = MediaPlanGenerator(self)
        generator.generate()
        self.close_waiting_message()

    def update_load(self):
        if not self.start_time_entry.get() or not self.end_time_entry.get():
            return

        ad_durations = []
        ad_repeats = []

        for ad_file_entry, ad_repeat_entry in zip(self.ad_files, self.ad_repeat_combo):
            ad_file = ad_file_entry.get()
            ad_dur = AudioSegment.from_file(ad_file).duration_seconds if os.path.exists(ad_file) else 0
            ad_repeat = ad_repeat_entry.get()

            try:
                ad_dur = int(ad_dur)
                ad_repeat = int(ad_repeat)
            except ValueError:
                return

            ad_durations.append(ad_dur)
            ad_repeats.append(ad_repeat)

        start_time = hour_to_seconds(self.start_time_entry.get())
        end_time = hour_to_seconds(self.end_time_entry.get())

        if end_time < start_time:
            working_time = 86400 + end_time - start_time
        else:
            working_time = end_time - start_time

        ad_sum = sum(dur * repeat for dur, repeat in zip(ad_durations, ad_repeats))
        percent = ad_sum / working_time
        self.load.config(text="Загружаемость: " + "{:.2f}".format(percent * 100) + "%")

    def save_json(self):
        data = {
            "music": self.music_files_entry.get(),
            "start": self.start_time_entry.get(),
            "end": self.end_time_entry.get()
        }

        if len(self.ad_frame_list) != 0:
            ads = []
            for i in range(len(self.ad_frame_list)):
                ad_info = []
                ad_info.append(self.ad_files[i].get())
                ad_info.append(self.ad_repeat_combo[i].get())
                ads.append(ad_info)
            data.update({"ad": ads})
        exe_path = os.path.abspath(sys.argv[0])
        exe_dir = os.path.dirname(exe_path)

        json_object = json.dumps(data, indent=4)
        with open(exe_dir + "/data.json", "w", encoding='ascii') as outfile:
            outfile.write(json_object)

        self.root.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = MediaPlanApp(root)
    root.protocol("WM_DELETE_WINDOW", app.save_json)
    root.mainloop()
