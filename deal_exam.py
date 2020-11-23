# coding=utf-8
import datetime
import re
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
    file_path = askopenfilename(filetypes=[('All Files', '*'),('XLS', '*.xls;*.xlsx')])
    pathFile.set(file_path)


def startIt():
    if len(pathFile.get()) <= 0:
        showinfo('提示', '先选择文件')
        return
    try:
        book = xlrd.open_workbook(pathFile.get())
        sh = book.sheet_by_index(0)
        exams = {}
        lines = []
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
            # lines.append('%s\t%s\n' % (row[2],row[3]))

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
                        # if sh.cell_type(rx,i) != 0:
                        #     lines.append('%s.%s\n' % (chr(i+61),row[i]))

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
            file = open("exams/%s.txt" % title, 'w', encoding='utf-8')
            types = ['单选题', '多选题', '判断题']
            for t in types:
                lines.append(t)
                lines.append('\n')
                ques = exam.get(t)
                for index, q in enumerate(ques):
                    lines.append('%d. %s %s\n' % (index+1, q['question'], q['answer']))
                    options = q['options']
                    for o in options:
                        v = o['v']
                        if isinstance(v, float):
                            lines.append('%s. %d\n' % (o['k'], o['v']))
                        elif isinstance(v, str):
                            if len(o['v'])>0:
                                lines.append('%s. %s\n' % (o['k'], o['v']))
                    lines.append('\n')

            file.writelines(lines)
            file.close()
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



def start_txt():
    if len(pathFile.get()) <= 0:
        showinfo('提示', '先选择文件')
        return
    try:
        with open(pathFile.get(), "r", encoding='utf-8') as f:
            try:
                exam = {'单选题': [], '多选题': [], '判断题': []}
                # 1 判断 2 单选 3 多选
                current_type = 0
                # ques_singles = []
                # ques_mutils = []
                # ques_rights = []
                current_questions = exam.get('判断题')
                question = None
                qss = []
                ass = []

                for line in f.readlines():
                    line = line.strip('\n')
                    # showlog(line)
                    if line.find('判断')>0:
                        current_questions = exam.get('判断题')
                        current_type = 3
                    elif line.find('单选')>0:
                        current_questions = exam.get('单选题')
                        current_type = 1
                    elif line.find('多选')>0:
                        current_questions = exam.get('多选题')
                        current_type = 2
                    else:
                        point_pos = line.find('.')
                        if point_pos>0 and point_pos<5:
                            if line[0].isdecimal():
                                if question != None:
                                    current_questions.append(question)
                                question = {"checked": False}
                                question["type"] = current_type
                                question["myAnswer"] = ''
                                question["options"] = []
                                a = ''
                                tmp = line[point_pos+1:].replace(' ','').replace('(','').replace(')','').replace('（','').replace('）','')
                                as_temps = ['A','B','C','D','E','F','G']
                                for as_temp in as_temps:
                                    if tmp.find(as_temp)>=0:
                                        a+=as_temp
                                        tmp = tmp.replace(as_temp,'')
                                q = tmp


                                # answer_pos = line.rfind('（')
                                # q = line[point_pos+1:answer_pos]
                                # a = line[answer_pos+1:len(line)-1]
                                question["question"] = q
                                question["answer"] = a
                                qss.append(q)
                                ass.append(a)
                                showlog(q+'\t'+a)
                                if current_type == 3:
                                    option = {}
                                    option['k'] = 'A'
                                    option['v'] = '正确'
                                    option['shouldcheck'] = question["answer"].find(option['k']) >= 0
                                    question["options"].append(option)
                                    option = {}
                                    option['k'] = 'B'
                                    option['v'] = '错误'
                                    option['shouldcheck'] = question["answer"].find(option['k']) >= 0
                                    question["options"].append(option)
                                pass
                            elif line[0].isalpha():
                                if question != None:
                                    if current_type == 1 or current_type == 2:
                                        option = {}
                                        option['k'] = line[0]
                                        option['v'] = line[2:]
                                        option['shouldcheck'] = question["answer"].find(option['k']) >= 0
                                        question["options"].append(option)
                                pass

                print('------------')
                print(qss)
                print(ass)
                with open("exams/%s.json" % 'title', "w", encoding='utf-8') as f:
                    json.dump(exam, f, ensure_ascii=False, sort_keys=True, indent=4)

                # data = f.readline()
                #     showlog(line)
            except Exception as e:
                showlog('出错了')
                print(e)
            f.close()
        # book = xlrd.open_workbook(pathFile.get())
        # sh = book.sheet_by_index(0)
        #

        showlog('已完成')

    except Exception as e:
        showlog('出错了')
        print(e)

    showlog('finished')


def start_txt_js():
    if len(pathFile.get()) <= 0:
        showinfo('提示', '先选择文件')
        return
    try:
        with open(pathFile.get(), "r", encoding='utf-8') as f:
            try:
                qss = []
                ass = []
                for line in f.readlines():
                    line = line.strip('\n')
                    if len(line)>0 and line[0].isdecimal():
                        a = ''
                        point_pos = line.find('.')
                        tmp = line[point_pos + 1:].replace(' ', '').replace('(', '').replace(')', '').replace('（',
                                                                                                              '').replace(
                            '）', '')
                        as_temps = ['A', 'B', 'C', 'D', 'E', 'F', 'G']
                        for as_temp in as_temps:
                            if tmp.find(as_temp) >= 0:
                                a += as_temp
                                tmp = tmp.replace(as_temp, '')
                        q = tmp
                        qss.append(q)
                        ass.append(a)

                print('------------')
                print(qss)
                print(ass)
                print(len(qss))
                print(len(ass))
            except Exception as e:
                showlog('出错了')
                print(e)
            finally:
                f.close()
        showlog('已完成')
    except Exception as e:
        showlog('出错了')
        print(e)


def export_json():
    if len(pathFile.get()) <= 0:
        showinfo('提示', '先选择文件')
        return
    try:
        with open(pathFile.get(), "r", encoding='utf-8') as load_f:
    # with open("exams/廉洁自律_纪律处分条例.json", 'r', encoding='utf-8') as load_f:
            load_dict = json.load(load_f)
            qss = []
            ass = []
            for questions in load_dict.values():
                for question in questions:
                    q = question['question']
                    q = re.sub(r'[^a-zA-Z\u4e00-\u9fa5\d]', "", q)
                    qss.append(q)
                    asss = []
                    answers = question['answer']
                    options = question['options']
                    for option in options:
                        if option['k'] in answers:
                            v = option['v']
                            if isinstance(v,str):
                                asss.append(v)
                            else:
                                asss.append('%d' % (v))
                    ass.append('￥'.join(asss))
            print(qss)
            print(ass)
    except Exception as e:
        showlog('出错了')
        showlog(str(e))
    pass


def main():
    logs.grid(row=0, column=0, rowspan=6)

    Button(root, text='选择题库文件', command=selectFile).grid(row=0, column=1)
    Entry(root, textvariable=pathFile).grid(row=1, column=1)
    Button(root, text='excel输出', command=startIt).grid(row=2, column=1, columnspan=1)
    Button(root, text='txt输出', command=start_txt).grid(row=3, column=1, columnspan=1)
    Button(root, text='json导出', command=export_json).grid(row=4, column=1, columnspan=1)
    Button(root, text='txt到js', command=start_txt_js).grid(row=5, column=1, columnspan=1)
    root.mainloop()


root = Tk()
logs = ScrolledText(root, width=40, height=30)
pathFile = StringVar()
if __name__ == '__main__':
    main()
