import os
import json
import random
import openpyxl
from datetime import datetime
from openpyxl.styles import PatternFill
from mutagen import File
import tkinter as tk
from tkinter import filedialog, messagebox


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
        self.root.title("Media Plan Generator")

        # Initialize variables
        self.ad_dur = []
        self.ad_repeat = []
        self.ad_name = []

        # Setup initial GUI elements
        self.setup_initial_gui()

    def setup_initial_gui(self):
        self.frame = tk.Frame(self.root, padx=10, pady=10)
        self.frame.pack(padx=10, pady=10)

        tk.Label(self.frame, text="Music Folder:", anchor='w').grid(row=0, column=0, sticky=tk.W, padx=5, pady=5)
        self.music_folder_entry = tk.Entry(self.frame, width=50)
        self.music_folder_entry.grid(row=0, column=1, padx=5, pady=5)
        tk.Button(self.frame, text="Browse", command=self.browse_music_folder).grid(row=0, column=2, padx=5, pady=5)

        tk.Label(self.frame, text="Objects File:", anchor='w').grid(row=1, column=0, sticky=tk.W, padx=5, pady=5)
        self.objects_file_entry = tk.Entry(self.frame, width=50)
        self.objects_file_entry.grid(row=1, column=1, padx=5, pady=5)
        tk.Button(self.frame, text="Browse", command=self.browse_objects_file).grid(row=1, column=2, padx=5, pady=5)

        tk.Label(self.frame, text="Ads Folder:", anchor='w').grid(row=2, column=0, sticky=tk.W, padx=5, pady=5)
        self.ads_folder_entry = tk.Entry(self.frame, width=50)
        self.ads_folder_entry.grid(row=2, column=1, padx=5, pady=5)
        tk.Button(self.frame, text="Browse", command=self.browse_ads_folder).grid(row=2, column=2, padx=5, pady=5)

        tk.Label(self.frame, text="Number of Advertisers:", anchor='w').grid(row=3, column=0, sticky=tk.W, padx=5,
                                                                             pady=5)
        self.ad_number_entry = tk.Entry(self.frame, width=10)
        self.ad_number_entry.grid(row=3, column=1, sticky=tk.W, padx=5, pady=5)

        tk.Button(self.frame, text="Next", command=self.setup_ad_inputs).grid(row=4, column=1, pady=10)

    def browse_music_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.music_folder_entry.delete(0, tk.END)
            self.music_folder_entry.insert(0, folder)

    def browse_objects_file(self):
        file = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if file:
            self.objects_file_entry.delete(0, tk.END)
            self.objects_file_entry.insert(0, file)

    def browse_ads_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.ads_folder_entry.delete(0, tk.END)
            self.ads_folder_entry.insert(0, folder)

    def setup_ad_inputs(self):
        try:
            self.ad_number = int(self.ad_number_entry.get())
        except ValueError:
            messagebox.showerror("Error", "Number of advertisers must be an integer.")
            return

        self.music_folder = self.music_folder_entry.get()
        self.objects_file = self.objects_file_entry.get()
        self.ads_folder = self.ads_folder_entry.get()

        if not self.music_folder or not self.objects_file or not self.ads_folder:
            messagebox.showerror("Error", "Please select all required folders and files.")
            return

        # Process ad data
        ad_count = 0
        for x in os.listdir(self.ads_folder):
            filename = os.fsdecode(x)
            self.ad_name.append(filename)
            ad_count += 1
            if ad_count == self.ad_number:
                break

        # Create the scrolling canvas and scrollbar
        self.ad_window = tk.Toplevel(self.root)
        self.ad_window.title("Enter Ad Durations and Repeats")

        self.canvas = tk.Canvas(self.ad_window, borderwidth=0)
        self.scroll_frame = tk.Frame(self.canvas)
        self.vsb = tk.Scrollbar(self.ad_window, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=self.vsb.set)

        self.vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.create_window((4, 4), window=self.scroll_frame, anchor="nw")

        self.scroll_frame.bind("<Configure>", lambda event, canvas=self.canvas: self.on_frame_configure(canvas))

        tk.Label(self.scroll_frame, text="Ad Name").grid(row=0, column=0, padx=5, pady=5)
        tk.Label(self.scroll_frame, text="Duration (seconds)").grid(row=0, column=1, padx=5, pady=5)
        tk.Label(self.scroll_frame, text="Repeats").grid(row=0, column=2, padx=5, pady=5)

        self.ad_entries = []

        for i in range(self.ad_number):
            ad_name_label = tk.Label(self.scroll_frame, text=self.ad_name[i])
            ad_name_label.grid(row=i + 1, column=0, padx=5, pady=5)

            ad_dur_entry = tk.Entry(self.scroll_frame)
            ad_dur_entry.grid(row=i + 1, column=1, padx=5, pady=5)

            ad_repeat_entry = tk.Entry(self.scroll_frame)
            ad_repeat_entry.grid(row=i + 1, column=2, padx=5, pady=5)

            self.ad_entries.append((ad_dur_entry, ad_repeat_entry))

        tk.Button(self.scroll_frame, text="Generate Media Plan", command=self.generate_media_plan).grid(
            row=self.ad_number + 1, column=1, pady=10)

    def on_frame_configure(self, canvas):
        '''Reset the scroll region to encompass the inner frame'''
        canvas.configure(scrollregion=canvas.bbox("all"))

    def generate_media_plan(self):
        self.ad_dur = []
        self.ad_repeat = []

        for ad_dur_entry, ad_repeat_entry in self.ad_entries:
            try:
                ad_dur = int(ad_dur_entry.get())
                ad_repeat = int(ad_repeat_entry.get())
            except ValueError:
                messagebox.showerror("Error", "Ad durations and repeats must be integers.")
                return
            self.ad_dur.append(ad_dur)
            self.ad_repeat.append(ad_repeat)

        current_datetime = datetime.now().strftime("%Y-%m-%d %H-%M")

        # Initialize variables
        song_name = []
        song_dur = []
        object_name = []
        object_time1 = []
        object_time2 = []

        # Process music files
        for x in os.listdir(self.music_folder):
            filename = os.fsdecode(x)
            song_name.append(filename)
            file_path = os.path.join(self.music_folder, filename)
            try:
                audio = File(file_path)
                if audio and hasattr(audio, 'info'):
                    duration = audio.info.length
                    song_dur.append(duration)
                else:
                    song_dur.append(None)
            except Exception as e:
                print(f"Error processing file {filename}: {e}")
                song_dur.append(None)

        # Process object data
        with open(self.objects_file, encoding='utf-8') as object_file:
            object_data = json.load(object_file)
            for x in object_data['objects']:
                object_name.append(x['Name'])
                object_time1.append(hour_to_seconds(x['time1']))
                object_time2.append(hour_to_seconds(x['time2']))

        indeces, self.ad_repeat = sort_list(self.ad_repeat)
        self.ad_name = rearrange(indeces, self.ad_name)
        self.ad_dur = rearrange(indeces, self.ad_dur)

        # Create workbook
        wb = openpyxl.Workbook()

        for times in range(len(object_name)):
            index_20_start = len(self.ad_name)
            for i in range(len(self.ad_name)):
                if self.ad_repeat[i] == 20:
                    index_20_start = i
                    break
            index_20_end = len(self.ad_name)
            index_20 = index_20_start

            if object_time2[times] < object_time1[times]:
                working_time = 86400 + object_time2[times] - object_time1[times]
            else:
                working_time = object_time2[times] - object_time1[times]

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

            blocks = block_period * 20

            ad_uses = []
            ad_broadcast = []
            for i in range(len(self.ad_repeat)):
                ad_uses.append(0)
                ad_broadcast.append(0)
            period = int((working_time - 20) / blocks)
            cur_time = object_time1[times]
            ad_start = object_time1[times] + period - 120
            cur_block = False
            ad_block = 1
            ads_in_block = 0
            timer = 0
            timestamps = []
            block_start = ''
            end_time = 0
            if object_time2[times] < object_time1[times]:
                end_time = 86400 + object_time2[times]
            else:
                end_time = object_time2[times]

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
            ws["H4"] = str(percent * 100) + '%'
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
                    if (ad_block) % block_period == 1:
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
                                ws['A' + str(row_excel)].fill = colors[cur_color]
                                ws['B' + str(row_excel)].fill = colors[cur_color]
                                row_excel += 1
                                ads_in_block += 1
                                ad_uses[i] += 1
                                timer += int(float(self.ad_dur[i])) + 1
                                cur_time += int(float(self.ad_dur[i])) + 1
                                continue

                        if self.ad_repeat[i] == 15 and ad_uses[i] != self.ad_repeat[i]:
                            if (ad_block % int(blocks / 15) == 1) or (ad_broadcast[i] == 1):
                                if ads_in_block == adlimit:
                                    ad_broadcast[i] = 1
                                    continue
                                ad_broadcast[i] = 0
                                ws["A" + str(row_excel)] = self.ad_name[i]
                                ws["B" + str(row_excel)] = sec_to_hour(cur_time)
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

        filename = f"Медиаплан {self.ad_number} реклам 20,15,10,5 повторов {sum(self.ad_dur)} сек {current_datetime}.xlsx"

        del wb['Sheet']
        wb.save(filename)
        messagebox.showinfo("Success", f"Media plan generated successfully and saved as {filename}!")


if __name__ == "__main__":
    root = tk.Tk()
    app = MediaPlanApp(root)
    root.mainloop()
