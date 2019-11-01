# coding=utf-8

import datetime
import copy
from tkinter import *
from tkinter.scrolledtext import ScrolledText
from tkinter.ttk import *
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import *

from datas import orgs,city_count
import xlsxwriter
import xlrd


def selectFile():
    file_path = askopenfilename(filetypes=[('XLS', '*.xls;*.xlsx'), ('All Files', '*')])
    pathFile.set(file_path)

def showlog(text):
    logs.insert(END, str(text) + "\n")
    logs.update()
    print(text)


def add_to_results(all_result,the_results,org):
    if not orgs.__contains__(org):
        return -1
    else:
        city = orgs.get(org)
        if not the_results.__contains__(city):
            the_results[city] = {}
        if not the_results[city].__contains__(org):
            the_results[city][org] = 0
        the_results[city][org]+=1
        all_result[city][org]+=1


def write_to_file(all_result,the_result):
    work_book = xlsxwriter.Workbook(pathFile.get() + '--' + datetime.datetime.now().strftime('%H%M') + '.xlsx')
    normal_format = work_book.add_format({
        'font_size': '12',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter', })
    worksheets = {}

    sheet_city = work_book.add_worksheet('按全部地市汇总')
    worksheets['按全部地市汇总'] = sheet_city
    wirte_results(sheet_city, normal_format,all_result)

    sheet_city = work_book.add_worksheet('按地市汇总')
    worksheets['按地市汇总'] = sheet_city
    wirte_results(sheet_city, normal_format,the_result)

    work_book.close()
    showlog('完成')


def wirte_results(sheet_summary, normal_format,the_result):
    m = 1
    for city,org_temp in the_result.items():
        size = len(org_temp)
        if size > 1:
            a = 'A%d:A%d' % (m + 1, m + size)
            b = 'B%d:B%d' % (m + 1, m + size)
            sheet_summary.merge_range(a, city, normal_format)
            sheet_summary.merge_range(b, '=SUM(D%d:D%d)' % (m + 1, m + size), normal_format)
        else:
            sheet_summary.write(m, 0, city, normal_format)
            sheet_summary.write(m, 1, '=SUM(D%d:D%d)' % (m + 1, m + size), normal_format)
        for org,org_count in org_temp.items():
            sheet_summary.write(m, 2, org, normal_format)
            sheet_summary.write(m, 3, org_count, normal_format)
            m = m + 1

    sheet_summary.write(m, 0, '合计', normal_format)
    sheet_summary.write(m, 1, '=SUM(D:D)', normal_format)


def startIt():
    if len(pathFile.get()) <= 0:
        showinfo('提示', '先选择文件')
        return
    book = xlrd.open_workbook(pathFile.get())
    if len(sheetName.get()) <= 0:
        sh = book.sheet_by_index(0)
    else:
        sh = book.sheet_by_name(sheetName)
    all_result = copy.deepcopy(city_count)
    the_result = {}
    error_count = 0
    col = start_col.get()-1
    for rx in range(start_row.get()-1, sh.nrows):
        row = (sh.row_values(rx))
        org = row[col].replace("农商", "").replace("商行", "").replace("银", "").replace("行", "")
        res = add_to_results(all_result,the_result,org)
        if res == -1:
            error_count += 1

    write_to_file(all_result,the_result)
    if error_count > 0:
        showlog('错误：'+error_count+'条')
    return


def main():
    logs.grid(row=0, column=0, rowspan=6)

    Button(root, text='选择Excel文件', command=selectFile).grid(row=0, column=1)
    Entry(root, textvariable=pathFile).grid(row=0, column=2)
    Label(root, text="sheet名称\n不填则取第一张", ).grid(row=1, column=1, columnspan=1)
    Entry(root, textvariable=sheetName).grid(row=1, column=2)

    Label(root, text="开始行号", ).grid(row=2, column=1, columnspan=1)
    Entry(root, textvariable=start_row).grid(row=2, column=2)

    Label(root, text="开始列号", ).grid(row=3, column=1, columnspan=1)
    Entry(root, textvariable=start_col).grid(row=3, column=2)

    Button(root, text='开始', command=startIt).grid(row=4, column=1)
    root.mainloop()


root = Tk()
logs = ScrolledText(root, width=40, height=30)
pathFile = StringVar()
sheetName = StringVar()
start_row = IntVar()
start_row.set(1)
start_col = IntVar()
start_col.set(1)
startline = -1
if __name__ == '__main__':
    main()
    # str = "Line1-abcdef \nLine2-abc \nLine4-abcd";
    # print(str.split('sdcx'))

