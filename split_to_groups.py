# coding=utf-8

import datetime
import copy
from tkinter import *
from tkinter.scrolledtext import ScrolledText
from tkinter.ttk import *
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import *

from datas import orgs, cities
import xlsxwriter
import xlrd


def selectFile():
    file_path = askopenfilename(filetypes=[('XLS', '*.xls;*.xlsx'), ('All Files', '*')])
    pathFile.set(file_path)


def showlog(text):
    logs.insert(END, str(text) + "\n")
    logs.update()
    print(text)


def add_to_results(all_result, org, row):
    if not orgs.__contains__(org):
        return -1
    else:
        city = orgs.get(org)
        all_result[city][org].append(row)


def write_to_file(all_result):
    work_book = xlsxwriter.Workbook(pathFile.get() + '--' + datetime.datetime.now().strftime('%H%M') + '.xlsx')
    normal_format = work_book.add_format({
        'font_size': '12',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter', })
    worksheets = {}

    sheet_city = work_book.add_worksheet('按全部地市汇总')
    worksheets['按全部地市汇总'] = sheet_city
    wirte_results(sheet_city, normal_format, all_result)

    work_book.close()
    showlog('完成')


def wirte_results(sheet_summary, normal_format, the_result):
    m = 1
    for i, results in enumerate(the_result):
    # for results in the_result:
        a = 'A%d:A%d' % (m+1, m + len(results))
        sheet_summary.merge_range(a, '第%d组' % (i+1), normal_format)
        for result in results:
            sheet_summary.write_row(m,1,result)
            m=m+1



def startIt():
    if len(pathFile.get()) <= 0:
        showinfo('提示', '先选择文件')
        return
    book = xlrd.open_workbook(pathFile.get())
    sh = book.sheet_by_index(0)
    error_count = 0
    all_result = []
    for i in range(0, group_count.get()):
        all_result.append([])

    col = start_col.get() - 1
    for rx in range(start_row.get() - 1, sh.nrows):
        row = (sh.row_values(rx))
        index = row[col]
        if isinstance(index, float):
            all_result[int(index-1) % group_count.get()].append(row)
        else:
            error_count += 1

    write_to_file(all_result)
    if error_count > 0:
        showlog('错误：' + error_count + '条')
    return


def main():
    logs.grid(row=0, column=0, rowspan=6)

    Button(root, text='选择Excel文件', command=selectFile).grid(row=0, column=1)
    Entry(root, textvariable=pathFile).grid(row=0, column=2)
    Label(root, text="分成几个组？", ).grid(row=1, column=1, columnspan=1)
    Entry(root, textvariable=group_count).grid(row=1, column=2)

    Label(root, text="开始行号", ).grid(row=2, column=1, columnspan=1)
    Entry(root, textvariable=start_row).grid(row=2, column=2)

    Label(root, text="序号在哪列", ).grid(row=3, column=1, columnspan=1)
    Entry(root, textvariable=start_col).grid(row=3, column=2)

    Button(root, text='开始', command=startIt).grid(row=4, column=1)
    root.mainloop()


root = Tk()
logs = ScrolledText(root, width=40, height=30)
pathFile = StringVar()
group_count = IntVar()
group_count.set(1)
start_row = IntVar()
start_row.set(1)
start_col = IntVar()
start_col.set(1)
startline = -1
if __name__ == '__main__':
    main()
    # str = "Line1-abcdef \nLine2-abc \nLine4-abcd";
    # print(str.split('sdcx'))
