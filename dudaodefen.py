# coding=utf-8
# 复制删除法按地市法人分割汇总文件
import datetime
import os
from pathlib import Path
from tkinter import *
from tkinter.scrolledtext import ScrolledText
from tkinter.ttk import *
from tkinter.filedialog import askopenfilename, askdirectory
from tkinter.messagebox import *
import openpyxl
import xlrd

import _thread

from utils import orgs


def select_folder():
    folder_path = askdirectory()
    pathFolder.set(folder_path)


def select_file():
    file_path = askopenfilename(filetypes=[('XLSX', '*.xlsx'), ('All Files', '*')])
    pathFile.set(file_path)


def show_log(text):
    logs.insert(END, str(text) + "\n")
    logs.update()
    print(text)


def start_work():
    # if len(pathFile.get()) <= 0:
    #     showinfo('提示', '先选择文件')
    #     return
    source_file = pathFile.get()
    source_folder = pathFolder.get()
    files = os.listdir(source_folder)
    print((files))
    datas = list()
    # datas = ['', 0, '']
    # datas = {}
    for file in files:
        full_path = os.path.join(source_folder, file)
        print(full_path)
        if os.path.isfile(full_path):
            result = get_xls_data(full_path)
            # for i in range(0,3):
            #     datas[i] = datas[i] + result[i]
            datas.append(result)
            # org, result = get_xls_data(full_path)
            # datas[org] = result
    result = list()
    for i in range(4, 17):
        result.append(['', 0, ''])
    for data in datas:
        for i in range(0, len(data)):
            result[i][0] += data[i][0]
            result[i][1] += data[i][1]
            result[i][2] += data[i][2]
    for i in range(0, len(result)):
        result[i][1] = result[i][1] / float(len(datas))
    print(result)
    write_city_file(result)
    return


def get_xls_data(full_path):
    # org = ''
    datas = list()
    try:
        book = xlrd.open_workbook(full_path)
        sh = book.sheet_by_index(0)
        org = sh.cell(2, 0).value.replace('被督导单位：', '')
        print(org)
        for i in range(4, 17):
            cell_a = sh.cell(i, 6)
            cell_b = sh.cell(i, 7)
            cell_c = sh.cell(i, 8)
            data = ['', 0, '']
            if cell_a.ctype == 1 and len(cell_a.value) > 0:
                data[0] = org + cell_a.value + '\n'
            elif cell_a.ctype == 2:
                data[0] = org + str(int(cell_a.value)) + '\n'
            if cell_b.ctype == 2:
                data[1] = int(cell_b.value)
            if cell_c.ctype == 1 and len(cell_c.value) > 0:
                data[2] = org + cell_c.value + '\n'
            elif cell_c.ctype == 2:
                data[2] = org + str(int(cell_c.value)) + '\n'
            # data[2] = data[2] + str(sh.cell(i, 8).value) + '\n'
            # data.append(org + sh.cell(i, 6).value)
            # data.append(sh.cell(i, 7).value)
            # data.append(org + sh.cell(i, 8).value)
            datas.append(data)
    except Exception as e:
        show_log('出错了')
        show_log(str(e))
    return datas


def write_city_file(datas):
    try:
        book = openpyxl.load_workbook('dudao_template.xlsx')
        sheet = book.active
        for i in range(0, len(datas)):
            sheet.cell(5 + i, 7).value = datas[i][0]
            sheet.cell(5 + i, 8).value = datas[i][1]
            sheet.cell(5 + i, 9).value = datas[i][2]

        source_folder = pathFolder.get()
        last_sp = source_folder.rfind('/')
        org = source_folder[last_sp + 1:]
        book.save('dudao_out/%s-%s农商银行督导打分表（2021-2季度）.xlsx' % (orgs[org],org))
        book.close()
        show_log('%s-%s已完成' % (orgs[org],org))
    except Exception as e:
        show_log(str(e))
    return


def begin_work():
    _thread.start_new_thread(start_work, ())


def main():
    logs.grid(row=0, column=0, rowspan=6)

    Button(root, text='选择文件夹', command=select_folder).grid(row=0, column=1)
    Entry(root, textvariable=pathFolder).grid(row=0, column=2)
    # Button(root, text='选择模板文件', command=select_file).grid(row=1, column=1)
    # Entry(root, textvariable=pathFile).grid(row=1, column=2)
    Label(root, text='汇总文件名称', ).grid(row=2, column=1)
    Entry(root, textvariable=target_file).grid(row=2, column=2)

    Button(root, text='开始', command=begin_work).grid(row=5, column=1, columnspan=1)
    Label(root, text='需要另存为xlsx格式\n需按摘取字段排序好\n空行过多需先清除内容后删除', ).grid(row=5, column=2, columnspan=1)
    root.mainloop()


root = Tk()
root.wm_title('督导得分汇总')
logs = ScrolledText(root, width=40, height=30)
pathFolder = StringVar()
pathFile = StringVar()
target_file = StringVar()
if __name__ == '__main__':
    main()
