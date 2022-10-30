from tkinter import *
from tkinter.filedialog import *
import re
import shutil
import xlrd
import xlwt
import os
import sys
from datetime import datetime, date, time

def License(actual_date):
    if int(actual_date) >= 20200101:
        print('License expired')
        sys.exit()
    else:
        return None

# Лицензия
actual_date = datetime.today().strftime("%Y%m%d") #получение системной даты
License(actual_date)

root = Tk()

def Quit(ev):
    global root
    root.destroy()


def create_txt(ev):

    op = askopenfilename()

    if op == '':
        return

    file = xlrd.open_workbook(op, formatting_info=True)

    folder = 'Output'
    if os.path.exists(folder):
        shutil.rmtree(folder)
    os.mkdir(folder)  # создание папки

    sheetsList = file.sheet_names()

    for j in range(len(sheetsList)):
        sheet = file.sheet_by_index(j)
        countsheet = 0
        for q in range(sheet.nrows):
            countsheet += 1

        if sheetsList[j] == 'Casing':
            countforrepeat = 0
            with open(r'Casing_out.tub', 'w', encoding='utf-8') as out:
                out.write('UNITS METRIC' + '\n' + '\n')

                for i in range(countsheet):
                    dateprev = sheet.row_values(i - 1)[1]
                    try:
                        datenext = sheet.row_values(i + 1)[1]
                        wellnext = sheet.row_values(i + 1)[0]
                    except:
                        pass

                    if i == 0 or i == 1:
                        continue
                    well = sheet.row_values(i)[0]
                    date = sheet.row_values(i)[1]
                    size = sheet.row_values(i)[2]
                    ot = sheet.row_values(i)[3]
                    do = sheet.row_values(i)[4]

                    if type(well) == float:
                        well = int(well)
                    elif type(well) == str:
                        well = str(well)

                    try:
                        y, m, d, h, i, s = xlrd.xldate_as_tuple(date, file.datemode)
                        date_perf_form = "{0}-{1}-{2}".format(y, m, d)
                    except:
                        date_perf_form = date
                    if dateprev == date:
                        countforrepeat += 1
                    if countforrepeat == 0:
                        out.write("DATE" + " " + str(date_perf_form) + "\n")
                        out.write("CASING" + ' ' + '"' + str(well) + '"' + ' ' + '"Casing ' + str(int(size)) + '"' + "\n")
                        out.write(str(int(ot)) + ' ' + '"' + str(int(size)) + '"' + "\n")
                        if datenext != date:#я тут
                            out.write(str(int(do)) + "\n" + "\n")
                    else:
                        if dateprev == date:
                            out.write(str(int(ot)) + ' ' + '"' + str(int(size)) + '"' + "\n")
                            if datenext != date:
                                out.write(str(int(do)) + "\n" + "\n")
                                countforrepeat = 0
                        else:
                            countforrepeat = 0
                out.close()
        elif sheetsList[j] == 'Tubing':
            countforrepeat = 0
            NameAdd = 0
            sizeAdd = 0
            with open(r'Tubing_out.tub', 'w', encoding='utf-8') as out:
                out.write('UNITS METRIC' + '\n' + '\n')

                for i in range(countsheet):
                    dateprev = sheet.row_values(i - 1)[1]
                    try:
                        datenext = sheet.row_values(i + 1)[1]
                    except:
                        pass

                    if i == 0 or i == 1:
                        continue
                    well = sheet.row_values(i)[0]
                    date = sheet.row_values(i)[1]
                    size = sheet.row_values(i)[2]
                    ot = sheet.row_values(i)[3]
                    do = sheet.row_values(i)[4]
                    wellprev = sheet.row_values(i-1)[0]
                    NameSizeAdd = 0

                    if NameAdd == 0:
                        NameSizeAdd = str(size)
                    else:
                        sizeAdd += 1
                        NameSizeAdd = str(size) + str(sizeAdd)
                    if well == wellprev:
                        if dateprev != date:
                            NameAdd += 1
                            sizeAdd += 1
                        else:
                            pass
                    else:
                        NameAdd = 0
                        sizeAdd = 0
                    if type(well) == float:
                        well = int(well)
                    elif type(well) == str:
                        well = str(well)

                    try:
                        y, m, d, h, i, s = xlrd.xldate_as_tuple(date, file.datemode)
                        date_perf_form = "{0}-{1}-{2}".format(y, m, d)
                    except:
                        date_perf_form = date
                    if dateprev == date:
                        countforrepeat += 1

                    if countforrepeat == 0:
                        out.write("DATE" + " " + str(date_perf_form) + "\n")
                        out.write("TUBING" + ' ' + '"Tubing ' + str(int(size)) + "_" + str(NameAdd)+ '"' + ' ' + '"' + str(well) + '"' + ' ' + '"' + str(well) + '"' + "\n")
                        out.write(str(int(ot)) + ' ' + '"' + str(int(size)) + '"' + "\n")

                        if datenext != date:
                            out.write(str(int(do)) + "\n" + "\n")
                    else:
                        if dateprev == date:
                            out.write(str(int(ot)) + ' ' + '"' + str(int(size)) + '"' + "\n")
                            if datenext != date:
                                out.write(str(int(do)) + "\n" + "\n")
                                countforrepeat = 0
                        else:
                            countforrepeat = 0


                out.close()


            """"
            count = 0
            with open(r'Tubing_out.tub', 'w', encoding='utf-8') as out:
                out.write('UNITS METRIC' + '\n' + '\n')

                for i in range(countsheet):
                    if i == 0 or i == 1:
                        continue
                    well = sheet.row_values(i)[0]
                    date = sheet.row_values(i)[1]
                    size = sheet.row_values(i)[2]
                    ot = sheet.row_values(i)[3]
                    do = sheet.row_values(i)[4]
                    if type(well) == float:
                        well = int(well)
                    elif type(well) == str:
                        well = str(well)

                    try:
                        y, m, d, h, i, s = xlrd.xldate_as_tuple(date, file.datemode)
                        date_perf_form = "{0}-{1}-{2}".format(y, m, d)
                    except:
                        date_perf_form = date

                    out.write("DATE" + " " + str(date_perf_form) + "\n")
                    out.write("TUBING" + ' ' + '"Tubing ' + str(int(size)) + '"' + ' ' + '"' + str(well) + '"' + ' ' + '"' + str(well) + '"' + "\n")
                    out.write(str(int(ot)) + ' ' + '"' + str(int(size)) + '"' + "\n")
                    out.write(str(int(do)) + "\n" + "\n")
                out.close()
                """""

        elif sheetsList[j] == 'Packers':
            count = 0
            with open(r'Packer_out.tub', 'w', encoding='utf-8') as out:
                out.write('UNITS METRIC' + '\n' + '\n')

                for i in range(countsheet):
                    if i == 0 or i == 1:
                        continue
                    well = sheet.row_values(i)[0]
                    date = sheet.row_values(i)[1]
                    ot = sheet.row_values(i)[2]
                    if type(well) == float:
                        well = int(well)
                    elif type(well) == str:
                        well = str(well)

                    try:
                        y, m, d, h, i, s = xlrd.xldate_as_tuple(date, file.datemode)
                        date_perf_form = "{0}-{1}-{2}".format(y, m, d)
                    except:
                        date_perf_form = date

                    out.write("DATE" + " " + str(date_perf_form) + "\n")
                    out.write("PACKER" + ' ' + '"Packer 1"' + " " + '"' + str(well) + '"' + ' ' + str(
                        int(ot)) + ' ' + '"PK_ADD_ON1"' + '\n' + '\n')
                out.close()

        elif sheetsList[j] == 'Perforations':
            count = 0
            with open(r'Perforation_out.ev', 'w', encoding='utf-8') as out:
                out.write('UNITS METRIC' + '\n')

                for i in range(countsheet):
                    if i == 1 or i == 0:
                        continue
                    else:
                        wellprev = sheet.row_values(i - 1)[0]
                    well = sheet.row_values(i)[0]
                    date = sheet.row_values(i)[1]
                    ot = sheet.row_values(i)[2]
                    do = sheet.row_values(i)[3]
                    if type(well) == float:
                        well = int(well)
                    elif type(well) == str:
                        well = str(well)

                    if well == wellprev:
                        count += 1
                    else:
                        count = 0
                        out.write('\n')
                    try:
                        y, m, d, h, i, s = xlrd.xldate_as_tuple(date, file.datemode)
                        date_perf_form = "{0}/{1}/{2}".format(d, m, y)
                    except:
                        date_perf_form = date

                    if count == 0:
                        out.write('WELLNAME' + ' ' + '"' + str(well) + '"' + '\n')
                    out.write(
                        str(date_perf_form) + ' ' + "perforation" + ' ' + str(int(ot)) + ' ' + str(int(do)) + '\n')
                out.close()

        elif sheetsList[j] == 'Plugs':
            count = 0
            with open(r'Plug_out.ev', 'w', encoding='utf-8') as out:
                out.write('UNITS METRIC' + '\n')

                for i in range(countsheet):
                    if i == 1 or i == 0:
                        continue
                    else:
                        wellprev = sheet.row_values(i - 1)[0]
                    well = sheet.row_values(i)[0]
                    date = sheet.row_values(i)[1]
                    ot = sheet.row_values(i)[2]

                    if type(well) == float:
                        well = int(well)
                    elif type(well) == str:
                        well = str(well)

                    if well == wellprev:
                        count += 1
                    else:
                        count = 0
                        out.write('\n')
                    try:
                        y, m, d, h, i, s = xlrd.xldate_as_tuple(date, file.datemode)
                        date_perf_form = "{0}/{1}/{2}".format(d, m, y)
                    except:
                        date_perf_form = date

                    if count == 0:
                        out.write('WELLNAME' + ' ' + '"' + str(well) + '"' + '\n')
                    out.write(str(date_perf_form) + ' ' + "plug" + ' ' + str(int(ot)) + '\n')
                out.close()

        elif sheetsList[j] == 'Stimulations':
            count = 0
            with open(r'Stimulation_out.ev', 'w', encoding='utf-8') as out:
                out.write('UNITS METRIC' + '\n')

                for i in range(countsheet):
                    if i == 1 or i == 0:
                        continue
                    else:
                        wellprev = sheet.row_values(i - 1)[0]
                    well = sheet.row_values(i)[0]
                    date = sheet.row_values(i)[1]
                    ot = sheet.row_values(i)[2]
                    do = sheet.row_values(i)[3]
                    if type(well) == float:
                        well = int(well)
                    elif type(well) == str:
                        well = str(well)

                    if well == wellprev:
                        count += 1
                    else:
                        count = 0
                        out.write('\n')
                    try:
                        y, m, d, h, i, s = xlrd.xldate_as_tuple(date, file.datemode)
                        date_perf_form = "{0}/{1}/{2}".format(d, m, y)
                    except:
                        date_perf_form = date

                    if count == 0:
                        out.write('WELLNAME' + ' ' + '"' + str(well) + '"' + '\n')
                    out.write(str(date_perf_form) + ' ' + "stimulate" + ' ' + str(int(ot)) + ' ' + str(int(do)) + '\n')
                out.close()

    shutil.move('Casing_out.tub', folder)
    shutil.move('Tubing_out.tub', folder)
    shutil.move('Packer_out.tub', folder)
    shutil.move('Perforation_out.ev', folder)
    shutil.move('Plug_out.ev', folder)
    shutil.move('Stimulation_out.ev', folder)
    Quit(ev)



panelFrame = Frame(root, height = 60, width = 100, bg = 'grey')      #Оболочка
textFrame = Frame(root, height = 0, width = 0)

panelFrame.pack(side='top', fill='x')
textFrame.pack(side='bottom', fill='both', expand=1)

textbox = Text(textFrame, font='Arial 14', wrap='word')
scrollbar = Scrollbar(textFrame)

scrollbar['command'] = textbox.yview
textbox['yscrollcommand'] = scrollbar.set

textbox.pack(side = 'left', fill = 'both', expand = 1)
scrollbar.pack(side = 'right', fill = 'y')

loadBtn = Button(panelFrame, text = 'Open file')
quitBtn = Button(panelFrame, text = 'Quit')

loadBtn.bind("<Button-1>", create_txt)
quitBtn.bind("<Button-1>", Quit)

loadBtn.place(x = 350, y = 10, width = 150, height = 40)
quitBtn.place(x = 800, y = 10, width = 40, height = 40)

root.mainloop()
