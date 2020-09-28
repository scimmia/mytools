# coding=utf-8
import datetime
from tkinter import *
from tkinter.scrolledtext import ScrolledText
from tkinter.ttk import *
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import *

import xlrd
import xlsxwriter
import json
# 处理题库

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
        exams = {}
        for rx in range(1,sh.nrows):
            row = (sh.row_values(rx))
            if not exams.__contains__(row[0]):
                exams[row[0]] = {'单选题':[],'多选题':[],'判断题':[]}
            exam = exams.get(row[0])
            q_type = 1
            # questions = exam.get(row[1])
            if row[1] == '单选题':
                q_type = 1
            elif row[1] == '多选题':
                q_type = 2
            elif row[1] == '判断题':
                q_type = 3
            question = {"checked": False}
            question["type"] = q_type
            question["question"] = row[2]
            question["answer"] = row[3]
            question["myAnswer"] = ''
            question["options"] = []
            if row[1] == '判断题':
                # question["options"] = {"A": "正确", "B": "错误",}
                option = {}
                option['k'] = 'A'
                option['v'] = '正确'
                option['shouldcheck'] = row[3].find(option['k']) >= 0
                question["options"].append(option)
                option = {}
                option['k'] = 'B'
                option['v'] = '错误'
                option['shouldcheck'] = row[3].find(option['k']) >= 0
                question["options"].append(option)
            else:
                for i in range(4,11):
                    option = {}
                    if sh.cell_type(rx, i) != 0:
                        option['k'] = chr(i+61)
                        option['v'] = row[i]
                        option['shouldcheck'] = row[3].find(option['k']) >= 0
                        question["options"].append(option)

                # question["options"] = {}
                # if sh.cell_type(rx,4) != 0:
                #     question["options"]['A'] = row[4]
                # if sh.cell_type(rx,5) != 0:
                #     question["options"]['B'] = row[5]
                # if sh.cell_type(rx,6) != 0:
                #     question["options"]['C'] = row[6]
                # if sh.cell_type(rx,7) != 0:
                #     question["options"]['D'] = row[7]
                # if sh.cell_type(rx,8) != 0:
                #     question["options"]['E'] = row[8]
                # if sh.cell_type(rx,9) != 0:
                #     question["options"]['F'] = row[9]
                # if sh.cell_type(rx,10) != 0:
                #     question["options"]['G'] = row[10]
            exam.get(row[1]).append(question)
        summary = []
        m = ''
        for title,exam in exams.items():
            s = {}
            s['title'] = title
            s['version'] = 1
            s['hidden'] = False
            s['s'] = len(exam.get('单选题'))
            s['m'] = len(exam.get('多选题'))
            s['r'] = len(exam.get('判断题'))
            s['total'] = s['s']+s['m']+s['r']
            s['cloudid'] = 'cloud://myexam-ofzc5.6d79-myexam-ofzc5-1254219446/exams/%s.json' % title
            m = m + json.dumps(s,ensure_ascii=False,sort_keys=True, indent=4) + '\n'
            summary.append(s)
            with open("exams/%s.json" % title, "w",encoding='utf-8') as f:
                json.dump(exam, f,ensure_ascii=False,sort_keys=True, indent=4)
        # showlog(m)
        file = open("exams/summary-%s.json" % datetime.datetime.now().strftime('%H%M'), 'w', encoding='utf-8')
        file.write(m)
        file.close()
        with open("exams/summary.json" , "w", encoding='utf-8') as f:
            json.dump(summary, f, ensure_ascii=False, sort_keys=True, indent=4)
        showlog('已完成')
    except Exception as e:
        showlog('出错了')
        showlog(str(e))

    showlog('finished')


def main():
    logs.grid(row=0, column=0, rowspan=6)

    Button(root, text='选择题库文件', command=selectFile).grid(row=0, column=1)
    Entry(root, textvariable=pathFile).grid(row=1, column=1)
    Button(root, text='输出', command=startIt).grid(row=2, column=1)
    root.mainloop()


root = Tk()
logs = ScrolledText(root, width=40, height=30)
pathFile = StringVar()
if __name__ == '__main__':
    main()
