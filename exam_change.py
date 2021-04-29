# coding=utf-8
import datetime
from tkinter import *
from tkinter.scrolledtext import ScrolledText
from tkinter.ttk import *
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import *

import xlrd
import xlsxwriter


def showlog(text):
    logs.insert(END, str(text) + "\n")
    logs.update()
    print(text)


def selectFile():
    file_path = askopenfilename(filetypes=[('XLS', '*.xls;*.xlsx'), ('All Files', '*')])
    pathFile.set(file_path)


def startIt():
    if len(pathFile.get()) <= 0:
        showinfo('提示', '先选择文件')
        return
    try:
        book = xlrd.open_workbook(pathFile.get())
        sh = book.sheet_by_index(0)

        ques_singles = []
        ques_mutils = []
        ques_rights = []
        for rx in range(sh.nrows):
            row = (sh.row_values(rx))
            if row[2] == '单选':
                ques_singles.append(row)
            elif row[2] == '多选':
                ques_mutils.append(row)
            elif row[2] == '判断':
                ques_rights.append(row)


        showlog('单选题')
        for s in ques_singles:
            showlog(s[0]+'. ' + s[1])
            if s[4].find('选项1') >= 0:
                showlog('\t'+s[5])
            elif s[4].find('选项2') >= 0:
                showlog('\t'+s[6])
            elif s[4].find('选项3') >= 0:
                showlog('\t'+s[7])
            elif s[4].find('选项4') >= 0:
                showlog('\t'+s[8])
        showlog('多选题')
        for s in ques_mutils:
            showlog(s[0]+'. ' + s[1])
            if s[4].find('选项1') >= 0:
                showlog('\t'+s[5])
            elif s[4].find('选项2') >= 0:
                showlog('\t'+s[6])
            elif s[4].find('选项3') >= 0:
                showlog('\t'+s[7])
            elif s[4].find('选项4') >= 0:
                showlog('\t'+s[8])
        showlog('判断题')
        for s in ques_rights:
            showlog(s[0]+'. ' + s[1])
            if s[4].find('选项1') >= 0:
                showlog('\t'+s[5])
            elif s[4].find('选项2') >= 0:
                showlog('\t'+s[6])
            elif s[4].find('选项3') >= 0:
                showlog('\t'+s[7])
            elif s[4].find('选项4') >= 0:
                showlog('\t'+s[8])
        showlog('已完成')
    except Exception as e:
        showlog('出错了')
        print(e)
    # except:
    #     showlog('出错了')

    showlog('finished')


def main():
    logs.grid(row=0, column=0, rowspan=6)

    Button(root, text='选择题库文件', command=selectFile).grid(row=0, column=1)
    Entry(root, textvariable=pathFile).grid(row=1, column=1)
    Button(root, text='excel输出', command=startIt).grid(row=2, column=1, columnspan=1)
    # Button(root, text='txt输出', command=start_txt).grid(row=3, column=1, columnspan=1)
    root.mainloop()


root = Tk()
logs = ScrolledText(root, width=40, height=30)
pathFile = StringVar()
if __name__ == '__main__':
    main()
