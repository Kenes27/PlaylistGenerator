import os
import random
import openpyxl
from datetime import datetime
from openpyxl.styles import PatternFill
from mutagen import File
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import *
from pydub import AudioSegment

def find_ab(percentage):
    if percentage == 0:
        return 0
    elif 0.05 >= percentage > 0:
        return 1
    elif 0.14 >= percentage > 0.05:
        return 2
    elif 0.19 >= percentage > 0.14:
        return 3
    elif 1 >= percentage > 0.14:
        return 4
    else:
        print("too much")
        return -1

def hour_to_seconds(time):
    a = 0
    b = ''
    index = 2
    for x in time:
        if ord(x) == ord(':'):
            a += int(b) * pow(60, index)
            index -= 1
            b = ''
            continue
        b += str(x)
    a += int(b)
    return a

def sec_to_hour(time):
    a = int(time / 3600)
    b = ''
    if len(str(a)) == 1:
        b = "0" + b + str(a) + ':'
    else:
        b = b + str(a) + ':'
    time -= a * 3600

    a = int(time / 60)
    if len(str(a)) == 1:
        b = b + '0' + str(a) + ':'
    else:
        b = b + str(a) + ':'

    time -= a * 60
    if len(str(time)) == 1:
        b = b + '0' + str(time)
    else:
        b = b + str(time)
    return b

def sort_list(repeat):
    pos = 0
    iter = 0
    arr_llst = []
    num = [5, 10, 15, 20]
    for i in range(len(repeat)):
        arr_llst.append(i)
    for j in num:
        for i in repeat[pos:len(repeat)]:
            if i == j:
                if pos == iter:
                    iter += 1
                    pos += 1
                    continue
                b = arr_llst[pos]
                arr_llst[pos] = arr_llst[iter]
                arr_llst[iter] = b
                a = repeat[pos]
                repeat[pos] = repeat[iter]
                repeat[iter] = a
                pos += 1
            iter += 1
        iter = pos
    return arr_llst, repeat

def rearrange(index, lists):
    new_list = []
    for i in range(len(index)):
        new_list.append(lists[index[i]])
    return new_list

class MediaPlanApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Генератор медиаплана")

        # Инициализация переменных
        self.ad_dur = []
        self.ad_repeat = []
        self.ad_name = []
        self.ad_files = []
        self.music_files = []

        # Настройка начального интерфейса
        self.setup_initial_gui()

    def setup_initial_gui(self):
        self.frame = tk.Frame(self.root, padx=10, pady=10)
        self.frame.pack(padx=10, pady=10)

        tk.Label(self.frame, text="Папка с музыкальные файлы:", anchor='w').grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.music_files_entry = tk.Entry(self.frame, width=50)
        self.music_files_entry.grid(row=0, column=1, padx=5, pady=5)
        tk.Button(self.frame, text="Обзор", command=self.browse_music_files).grid(row=0, column=2, padx=5, pady=5)

        tk.Label(self.frame, text="Начальное время:", anchor='w').grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.start_time_entry = tk.Entry(self.frame, width=10)
        self.start_time_entry.grid(row=1, column=1, sticky=tk.W, padx=5, pady=5)

        tk.Label(self.frame, text="Конечное время:", anchor='w').grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.end_time_entry = tk.Entry(self.frame, width=10)
        self.end_time_entry.grid(row=2, column=1, sticky=tk.W, padx=5, pady=5)

        tk.Button(self.frame, text="Добавить рекламу", command=self.add_advertisement).grid(row=3, column=1, pady=10)

        # Create a canvas for scrolling
        self.canvas = Canvas(self.root, borderwidth=0, width=500, height=300)
        self.ad_frame = tk.Frame(self.canvas, padx=10, pady=10)
        self.scrollbar = Scrollbar(self.root, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.create_window((4, 4), window=self.ad_frame, anchor="nw", tags="self.ad_frame")

        self.ad_frame.bind("<Configure>", lambda event, canvas=self.canvas: self.on_frame_configure(canvas))

        tk.Button(self.ad_frame, text="Сгенерировать медиаплан", command=self.generate_media_plan).pack(pady=10)

    def on_frame_configure(self, canvas):
        canvas.configure(scrollregion=canvas.bbox("all"))

    def browse_music_files(self):
        files = filedialog.askdirectory()
        if files:
            self.music_files_entry.delete(0, tk.END)
            self.music_files_entry.insert(0, files)

    def add_advertisement(self):
        ad_frame = tk.Frame(self.ad_frame)
        ad_frame.pack(padx=5, pady=5)

        tk.Label(ad_frame, text="Файл рекламы:").grid(row=0, column=0, padx=5, pady=5)
        ad_file_entry = tk.Entry(ad_frame, width=30)
        ad_file_entry.grid(row=0, column=1, padx=5, pady=5)
        tk.Button(ad_frame, text="Обзор", command=lambda: self.browse_ad_file(ad_file_entry)).grid(row=0, column=2, padx=5, pady=5)

        #tk.Label(ad_frame, text="Продолжительность (секунды):").grid(row=1, column=0, padx=5, pady=5)
        #ad_dur_entry = tk.Entry(ad_frame, width=10)
        #ad_dur_entry.grid(row=1, column=1, padx=5, pady=5)

        tk.Label(ad_frame, text="Повторы:").grid(row=2, column=0, padx=5, pady=5)
        ad_repeat_entry = ttk.Combobox(ad_frame, values=["20", "15", "10", "5"])
        ad_repeat_entry.set("20")
        ad_repeat_entry.grid(row=2, column=1, padx=5, pady=5)

        self.ad_files.append(ad_file_entry)
        #self.ad_dur.append(ad_dur_entry)
        self.ad_repeat.append(ad_repeat_entry)

    def browse_ad_file(self, ad_file_entry):
        file = filedialog.askopenfilename(filetypes=[("Audio files", "*.mp3 *.wav *.flac *.m4a *.ogg")])
        if file:
            ad_file_entry.delete(0, tk.END)
            ad_file_entry.insert(0, file)

    def generate_media_plan(self):
        ad_durations = []
        ad_repeats = []
        ad_names = []

        for ad_file_entry, ad_repeat_entry in zip(self.ad_files, self.ad_repeat):
            ad_file = ad_file_entry.get()
            ad_name = os.path.basename(ad_file)
            ad_dur = AudioSegment.from_file(ad_file).duration_seconds
            ad_repeat = ad_repeat_entry.get()

            if not ad_file or not ad_dur or not ad_repeat:
                messagebox.showerror("Ошибка", "Пожалуйста, заполните все поля для каждой рекламы.")
                return

            try:
                ad_dur = int(ad_dur)
                ad_repeat = int(ad_repeat)
            except ValueError:
                messagebox.showerror("Ошибка", "Продолжительность и количество повторов для рекламы должны быть целыми числами.")
                return

            ad_durations.append(ad_dur)
            ad_repeats.append(ad_repeat)
            ad_names.append(ad_name)

        self.ad_dur = ad_durations
        self.ad_repeat = ad_repeats
        self.ad_name = ad_names

        self.start_time = self.start_time_entry.get()
        self.end_time = self.end_time_entry.get()

        if not self.start_time or not self.end_time:
            messagebox.showerror("Ошибка", "Пожалуйста, введите начальное и конечное время.")
            return

        current_datetime = datetime.now().strftime("%Y-%m-%d %H-%M")

        # Инициализация переменных
        song_name = []
        song_dur = []

        # Обработка музыкальных файлов
        
        for file_path in os.listdir(self.music_files_entry.get()):
            filename = os.path.basename(file_path)
            try:
                song = AudioSegment.from_file(self.music_files_entry.get() + '/'+ file_path).duration_seconds
                song_dur.append(song)
                song_name.append(filename)
            except Exception as e:
                print(f"Ошибка при обработке файла {filename}: {e}")
                #song_dur.append(None)

        # Преобразование времени
        object_time1 = hour_to_seconds(self.start_time)
        object_time2 = hour_to_seconds(self.end_time)

        indeces, self.ad_repeat = sort_list(self.ad_repeat)
        self.ad_name = rearrange(indeces, self.ad_name)
        self.ad_dur = rearrange(indeces, self.ad_dur)

        # Создание рабочей книги
        wb = openpyxl.Workbook()

        index_20_start = len(self.ad_name)
        for i in range(len(self.ad_name)):
            if self.ad_repeat[i] == 20:
                index_20_start = i
                break
        index_20_end = len(self.ad_name)
        index_20 = index_20_start

        if object_time2 < object_time1:
            working_time = 86400 + object_time2 - object_time1
        else:
            working_time = object_time2 - object_time1

        ad_sum = 0
        for k in range(len(self.ad_name)):
            ad_sum += self.ad_dur[k] * self.ad_repeat[k]
            percent = ad_sum / working_time
            adlimit = find_ab(percent)
        block_period = 0
        if len(self.ad_name) % adlimit == 0:
            block_period = len(self.ad_name) / adlimit
        else:
            block_period = int(len(self.ad_name) / adlimit) + 1

        blocks = int(block_period * 20)

        ad_uses = []
        ad_broadcast = []
        for i in range(len(self.ad_repeat)):
            ad_uses.append(0)
            ad_broadcast.append(0)
        period = int((working_time - 20) / blocks)
        cur_time = object_time1
        ad_start = object_time1 + period - 120
        cur_block = False
        ad_block = 1
        ads_in_block = 0
        timer = 0
        timestamps = []
        block_start = ''
        end_time = 0
        if object_time2 < object_time1:
            end_time = 86400 + object_time2
        else:
            end_time = object_time2

        wb.create_sheet(str(int(working_time / 3600)))
        ws = wb[str(int(working_time / 3600))]
        ws["A1"] = "Имя"
        ws['C1'] = "Время"
        ws["D1"] = str(int(working_time / 3600)) + ' часа'
        ws["G1"] = 'Реклама'
        ws["H1"] = len(self.ad_name)
        ws["G2"] = "Повторы"
        ws["H2"] = "20:15:10:5"
        ws["G3"] = 'Продолжительность'
        ws["H3"] = sum(self.ad_dur)
        ws["G4"] = "Загруженность"
        ws["H4"] = "{:.2f}".format(percent * 100) + '%'
        ws["G5"] = "Максимальный рекламный блок"
        ws["H5"] = adlimit

        row_excel = 2
        colors = []
        green = PatternFill(patternType='solid', fgColor='78D542')
        colors.append(green)
        yellow = PatternFill(patternType='solid', fgColor='FFC638')
        colors.append(yellow)
        cur_color = 1

        while cur_time < end_time:
            block_start = cur_time
            while cur_block == False:
                cur_color = 1
                if cur_time > ad_start:
                    if int(ad_start) - block_start < 0:
                        col = "C"
                        while ws[col + str(row_excel - 1)].value is not None:
                            col = chr(ord(col) + 1)
                        ws[col + str(row_excel - 1)] = "00:00:00"
                        ws[chr(ord(col) + 1) + str(row_excel - 1)] = "Музыка"
                    else:
                        cur_time = ad_start
                        col = "C"
                        while ws[col + str(row_excel - 1)].value is not None:
                            col = chr(ord(col) + 1)
                        ws[col + str(row_excel - 1)] = sec_to_hour(int(cur_time) - block_start)
                        ws[chr(ord(col) + 1) + str(row_excel - 1)] = "Музыка"
                    ad_start += period
                    cur_block = True
                    continue
                rand_int = random.randint(0, len(song_name) - 1)
                ws["A" + str(row_excel)] = song_name[rand_int]
                ws["B" + str(row_excel)] = sec_to_hour(cur_time)

                ws['A' + str(row_excel)].fill = colors[cur_color]
                ws['B' + str(row_excel)].fill = colors[cur_color]
                row_excel += 1
                cur_time += int(float(song_dur[rand_int])) + 1

            while cur_block == True:
                cur_color = 0
                if (ad_block - 1) == blocks:
                    cur_block = False
                    continue
                timer = 0
                block_start = cur_time
                reset = 1
                if block_period == 1:
                    reset = 0
                if (ad_block) % block_period == reset:
                    index_20 = index_20_start
                for i in range(len(self.ad_repeat)):
                    if self.ad_repeat[i] == 5 and ad_uses[i] != self.ad_repeat[i]:
                        if (ad_block % int(blocks / 5) == 1) or (ad_broadcast[i] == 1):
                            if ads_in_block == adlimit:
                                ad_broadcast[i] = 1
                                continue
                            ad_broadcast[i] = 0
                            ws["A" + str(row_excel)] = self.ad_name[i]
                            ws["B" + str(row_excel)] = sec_to_hour(cur_time)
                            ws["C" + str(row_excel)] = ad_uses[i] + 1
                            ws['A' + str(row_excel)].fill = colors[cur_color]
                            ws['B' + str(row_excel)].fill = colors[cur_color]
                            row_excel += 1
                            ads_in_block += 1
                            ad_uses[i] += 1
                            timer += int(float(self.ad_dur[i])) + 1
                            cur_time += int(float(self.ad_dur[i])) + 1
                            continue

                    if self.ad_repeat[i] == 10 and ad_uses[i] != self.ad_repeat[i]:
                        if (ad_block % int(blocks / 10) == 1) or (ad_broadcast[i] == 1):
                            if ads_in_block == adlimit:
                                ad_broadcast[i] = 1
                                continue
                            ad_broadcast[i] = 0
                            ws["A" + str(row_excel)] = self.ad_name[i]
                            ws["B" + str(row_excel)] = sec_to_hour(cur_time)
                            ws["C" + str(row_excel)] = ad_uses[i] + 1
                            ws['A' + str(row_excel)].fill = colors[cur_color]
                            ws['B' + str(row_excel)].fill = colors[cur_color]
                            row_excel += 1
                            ads_in_block += 1
                            ad_uses[i] += 1
                            timer += int(float(self.ad_dur[i])) + 1
                            cur_time += int(float(self.ad_dur[i])) + 1
                            continue

                    if self.ad_repeat[i] == 15 and ad_uses[i] != self.ad_repeat[i]:
                        mod = 1
                        if int(blocks / 15) == 1:
                            mod = 0
                        if (ad_block % int(blocks / 15) == mod) or (ad_broadcast[i] == 1):
                            if ads_in_block == adlimit:
                                ad_broadcast[i] = 1
                                continue
                            ad_broadcast[i] = 0
                            ws["A" + str(row_excel)] = self.ad_name[i]
                            ws["B" + str(row_excel)] = sec_to_hour(cur_time)
                            ws["C" + str(row_excel)] = ad_uses[i] + 1
                            ws['A' + str(row_excel)].fill = colors[cur_color]
                            ws['B' + str(row_excel)].fill = colors[cur_color]
                            row_excel += 1
                            ads_in_block += 1
                            ad_uses[i] += 1
                            timer += int(float(self.ad_dur[i])) + 1
                            cur_time += int(float(self.ad_dur[i])) + 1
                            continue

                    if self.ad_repeat[i] == 20 and ad_uses[i] != self.ad_repeat[i]:
                        if (ads_in_block == adlimit or i != index_20):
                            continue
                        ws["A" + str(row_excel)] = self.ad_name[i]
                        ws["B" + str(row_excel)] = sec_to_hour(cur_time)
                        ws["C" + str(row_excel)] = ad_uses[i] + 1
                        ws['A' + str(row_excel)].fill = colors[cur_color]
                        ws['B' + str(row_excel)].fill = colors[cur_color]
                        row_excel += 1
                        ads_in_block += 1
                        ad_uses[i] += 1
                        index_20 += 1
                        timer += int(float(self.ad_dur[i])) + 1
                        cur_time += int(float(self.ad_dur[i])) + 1
                        continue

                col = "C"
                while ws[col + str(row_excel - 1)].value is not None:
                    col = chr(ord(col) + 1)
                ws[col + str(row_excel - 1)] = sec_to_hour(int(cur_time) - block_start)
                ws[chr(ord(col) + 1) + str(row_excel - 1)] = "Реклама"
                ads_in_block = 0
                ad_block += 1
                cur_block = False

        ws.append([''])
        ws.append([''])
        ws.append(['Повторов за день'])
        for row in zip(ad_uses, self.ad_name):
            ws.append(row)

        ws.column_dimensions['A'].width = 66

        filename = f"Медиаплан {len(self.ad_name)} реклам 20,15,10,5 повторов {sum(self.ad_dur)} сек {current_datetime}.xlsx"

        del wb['Sheet']
        wb.save(filename)
        messagebox.showinfo("Успех", f"Медиаплан успешно создан и сохранен под именем {filename}!")


if __name__ == "__main__":
    root = tk.Tk()
    app = MediaPlanApp(root)
    root.mainloop()
