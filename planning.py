import os
import json
import random
import openpyxl
from datetime import datetime
from openpyxl.styles import PatternFill
from mutagen import File

# Move to the correct directory
os.chdir(r"File")
current_datetime = datetime.now().strftime("%Y-%m-%d %H-%M")

# Function to find the ad block


def find_ab(percentage):
    global adblock

    if percentage == 0:
        adblock = 0
    elif 0.05 >= percentage > 0:
        adblock = 1
    elif 0.14 >= percentage > 0.05:
        adblock = 2
    elif 0.19 >= percentage > 0.14:
        adblock = 3
    elif 1 >= percentage > 0.14:
        adblock = 4
    else:
        print("too much")

    return adblock


# Convert hour format to seconds
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


# Convert seconds to hour format
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


# Sort the list
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


# Rearrange the list based on indices
def rearrange(index, lists):
    new_list = []
    for i in range(len(index)):
        new_list.append(lists[index[i]])
    return new_list


# Initialize variables
song_name = []
song_dur = []
object_name = []
object_time1 = []
object_time2 = []
ad_name = []
ad_dur = []
ad_repeat = []

# Process music files
music = r'music'
for x in os.listdir(music):
    filename = os.fsdecode(x)
    song_name.append(filename)
    file_path = os.path.join(music, filename)
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
object_file = open(r'objects.json', encoding='utf-8')
object_data = json.load(object_file)
for x in object_data['objects']:
    object_name.append(x['Name'])
    object_time1.append(hour_to_seconds(x['time1']))
    object_time2.append(hour_to_seconds(x['time2']))

# Process ad data
ad_number = 14
ad_count = 0
ad = r'ad'
for x in os.listdir(ad):
    filename = os.fsdecode(x)
    ad_name.append(filename)
    ad_count += 1
    if ad_count == ad_number:
        break

all_dur = 30
ad_dur = [30, 26, 39, 24, 29, 36, 33, 20, 36, 34, 22, 28, 31, 34, 29, 26, 27, 19, 35, 25, 31, 32, 26, 28, 30, 25, 29,
          27, 34, 36, 37, 34, 38]

ad_repeat = [15, 20, 20, 20, 10, 20, 15, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 15, 5, 20, 15, 20, 20, 5, 20, 20,
             20, 20, 20, 15, 5, 20]

ad_repeat = ad_repeat[0:ad_number]

indeces, ad_repeat = sort_list(ad_repeat)
ad_name = rearrange(indeces, ad_name)
ad_dur = rearrange(indeces, ad_dur)

# Create workbook
wb = openpyxl.Workbook()

for times in range(len(object_name)):
    index_20_start = len(ad_name)
    for i in range(len(ad_name)):
        if ad_repeat[i] == 20:
            index_20_start = i
            break
    index_20_end = len(ad_name)
    index_20 = index_20_start

    if object_time2[times] < object_time1[times]:
        working_time = 86400 + object_time2[times] - object_time1[times]
    else:
        working_time = object_time2[times] - object_time1[times]

    ad_sum = 0
    for k in range(len(ad_name)):
        ad_sum += ad_dur[k] * ad_repeat[k]
        percent = ad_sum / working_time
        adlimit = find_ab(percent)
    block_period = 0
    if len(ad_name) % adlimit == 0:
        block_period = len(ad_name) / adlimit
    else:
        block_period = int(len(ad_name) / adlimit) + 1

    blocks = block_period * 20

    ad_uses = []
    ad_broadcast = []
    for i in range(len(ad_repeat)):
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
    ws["H1"] = len(ad_name)
    ws["G2"] = "Повторы"
    ws["H2"] = "20:15:10:5"
    ws["G3"] = 'Продолжительность'
    ws["H3"] = all_dur
    ws["G4"] = "Загруженность"
    ws["H4"] = str(percent * 100) + '%'
    ws["G5"] = "Максимальный рекламный блок"
    ws["H5"] = adlimit

    row_excel = 2
    colors = []
    green = PatternFill(patternType='solid',
                        fgColor='78D542')
    colors.append(green)
    yellow = PatternFill(patternType='solid',
                         fgColor='FFC638')
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
            for i in range(len(ad_repeat)):
                if ad_repeat[i] == 5 and ad_uses[i] != ad_repeat[i]:
                    if (ad_block % int(blocks / 5) == 1) or (ad_broadcast[i] == 1):
                        if ads_in_block == adlimit:
                            ad_broadcast[i] = 1
                            continue
                        ad_broadcast[i] = 0
                        ws["A" + str(row_excel)] = ad_name[i]
                        ws["B" + str(row_excel)] = sec_to_hour(cur_time)
                        ws['A' + str(row_excel)].fill = colors[cur_color]
                        ws['B' + str(row_excel)].fill = colors[cur_color]
                        row_excel += 1
                        ads_in_block += 1
                        ad_uses[i] += 1
                        timer += int(float(ad_dur[i])) + 1
                        cur_time += int(float(ad_dur[i])) + 1
                        continue

                if ad_repeat[i] == 10 and ad_uses[i] != ad_repeat[i]:
                    if (ad_block % int(blocks / 10) == 1) or (ad_broadcast[i] == 1):
                        if ads_in_block == adlimit:
                            ad_broadcast[i] = 1
                            continue
                        ad_broadcast[i] = 0
                        ws["A" + str(row_excel)] = ad_name[i]
                        ws["B" + str(row_excel)] = sec_to_hour(cur_time)
                        ws['A' + str(row_excel)].fill = colors[cur_color]
                        ws['B' + str(row_excel)].fill = colors[cur_color]
                        row_excel += 1
                        ads_in_block += 1
                        ad_uses[i] += 1
                        timer += int(float(ad_dur[i])) + 1
                        cur_time += int(float(ad_dur[i])) + 1
                        continue

                if ad_repeat[i] == 15 and ad_uses[i] != ad_repeat[i]:
                    if (ad_block % int(blocks / 15) == 1) or (ad_broadcast[i] == 1):
                        if ads_in_block == adlimit:
                            ad_broadcast[i] = 1
                            continue
                        ad_broadcast[i] = 0
                        ws["A" + str(row_excel)] = ad_name[i]
                        ws["B" + str(row_excel)] = sec_to_hour(cur_time)
                        ws['A' + str(row_excel)].fill = colors[cur_color]
                        ws['B' + str(row_excel)].fill = colors[cur_color]
                        row_excel += 1
                        ads_in_block += 1
                        ad_uses[i] += 1
                        timer += int(float(ad_dur[i])) + 1
                        cur_time += int(float(ad_dur[i])) + 1
                        continue

                if ad_repeat[i] == 20 and ad_uses[i] != ad_repeat[i]:
                    if (ads_in_block == adlimit or i != index_20):
                        continue
                    ws["A" + str(row_excel)] = ad_name[i]
                    ws["B" + str(row_excel)] = sec_to_hour(cur_time)
                    ws['A' + str(row_excel)].fill = colors[cur_color]
                    ws['B' + str(row_excel)].fill = colors[cur_color]
                    row_excel += 1
                    ads_in_block += 1
                    ad_uses[i] += 1
                    index_20 += 1
                    timer += int(float(ad_dur[i])) + 1
                    cur_time += int(float(ad_dur[i])) + 1
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
    for row in zip(ad_uses, ad_name):
        ws.append(row)

    ws.column_dimensions['A'].width = 66

filename = f"Медиаплан {ad_number} реклам 20,15,10,5 повторов {all_dur} сек {current_datetime}.xlsx"

del wb['Sheet']
wb.save(filename)
