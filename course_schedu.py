# coding=utf-8
import datetime
from tkinter import *
from tkinter.scrolledtext import ScrolledText
from tkinter.ttk import *
from tkinter.messagebox import *

import xlrd
import xlsxwriter
from collections import OrderedDict
from datas import format_header, format_merge, format_normal
import utils
import copy
import numpy as np

sheet_name_class_simple = '排课简略'
sheet_name_class_all = '排课总览'
sheet_name_class_middle = '排课安排'
sheet_name_teacher_info = '讲师'
pre_count = 0


class Course(object):
    def __init__(self, c_date, c_time, c_class, teacher):
        self.c_date = c_date
        self.c_time = c_time
        self.c_class = c_class
        self.c_teacher = teacher

    def get_data(self):
        return [self.c_date, self.c_time, self.c_teacher.c_course_name, '内部师资', self.c_teacher.c_part,
                self.c_teacher.c_teacher_name, ]


class Teacher(object):
    def __init__(self, c_teacher_name, c_course_name, c_part, c_from):
        self.c_teacher_name = c_teacher_name
        self.c_course_name = c_course_name
        self.c_part = c_part
        self.c_from = c_from

    def get_data(self):
        return [self.c_teacher_name, self.c_course_name, self.c_part, self.c_from]

    def get_class_data(self):
        return [self.c_course_name, self.c_from, self.c_part, self.c_teacher_name]


def showlog(text):
    logs.insert(END, str(text) + "\n")
    logs.update()
    print(text)


def make_mid_file():
    if len(pathFile.get()) <= 0:
        showinfo('提示', '先选择课程文件')
        return
    try:
        book = xlrd.open_workbook(pathFile.get())
        sh = book.sheet_by_index(0)
        col_date = 0
        col_time = 1
        col_class = 2
        col_from = 3
        col_org = 4
        col_teacher = 5
        dates = {}
        teacher_classes = {}
        the_date = 'default'
        found_header = False
        for rx in range(0, sh.nrows):
            row = sh.row(rx)
            if not found_header:
                if row[col_date].ctype == 1 and row[col_date].value.find('时间') >= 0:
                    found_header = True
            else:
                if row[col_date].ctype == 1:
                    if len(row[col_date].value) > 0:
                        the_date = row[col_date].value
                elif row[col_date].ctype == 3:
                    temp = xlrd.xldate.xldate_as_datetime(row[col_date].value, 0)
                    the_date = temp.strftime('%m-%d')
                    # print(the_date)
                if not dates.__contains__(the_date):
                    dates[the_date] = []
                dates[the_date].append(row[col_time].value)
                teacher_name = row[col_teacher].value
                if len(teacher_name) <= 0:
                    teacher_name = '无'
                if teacher_classes.__contains__(teacher_name):
                    teacher_name += str(rx)
                teacher_classes[teacher_name] = Teacher(teacher_name, row[col_class].value, row[col_org].value,
                                                        row[col_from].value)
        export_mid_file(dates, teacher_classes)
        showlog('已完成')
    except Exception as e:
        showlog('出错了')
        showlog(str(e))
    showlog('')


def export_mid_file(dates, teacher_classes):
    file_name = pathFile.get() + '-中间表' + datetime.datetime.now().strftime('%H%M') + '.xlsx'
    work_book = xlsxwriter.Workbook(file_name)
    head_format = work_book.add_format(format_header)
    merge_format = work_book.add_format(format_merge)
    normal_format = work_book.add_format(format_normal)

    ws = work_book.add_worksheet(sheet_name_class_simple)
    work_book.add_worksheet(sheet_name_class_middle)
    write_dates_to_sheet(ws, dates, merge_format)
    ws.write_row(0, 2, classes.get().split(" "), head_format)

    wsa = work_book.add_worksheet(sheet_name_class_all + 'A')
    write_dates_to_sheet(wsa, dates, merge_format)
    wsa.write_row(0, 2, teacher_classes.keys(), normal_format)
    wsa.write_row(0, 2 + len(teacher_classes.keys()), classes.get().split(" "), head_format)

    wsb = work_book.add_worksheet(sheet_name_class_all)
    write_dates_to_sheet(wsb, dates, merge_format)
    wsb.write_row(0, 2, teacher_classes.keys(), normal_format)
    wsb.write_row(0, 2 + len(teacher_classes.keys()), classes.get().split(" "), head_format)

    write_teacher_info_sheet(work_book, teacher_classes, normal_format)

    work_book.close()
    path_course_file.set(file_name)
    showlog('中间表-，文件为：')
    showlog(file_name)


# # 生成排课表
# def export_class_file(date_list, class_courses, teacher_class):
#     work_book = xlsxwriter.Workbook(path_course_file.get() + datetime.datetime.now().strftime('%H%M') + '排课表.xlsx')
#     head_format = work_book.add_format(format_header)
#     merge_format = work_book.add_format(format_merge)
#     normal_format = work_book.add_format(format_normal)
#
#     for name, class_course in class_courses.items():
#         sheet_name = name
#         ws = work_book.add_worksheet(sheet_name)
#         ws.default_row_height = 20
#         ws.set_column(0, 1, 15)
#         ws.set_column(2, 2, 60)
#         ws.set_column(3, 3, 10)
#         ws.set_column(4, 4, 15)
#         ws.set_column(5, 6, 10)
#         ws.set_row(0, 50)
#         ws.merge_range(0, 0, 0, 6, ('经营网点转型发展业务骨干培训班课程表--%s' % name), head_format)
#         ws.merge_range(1, 0, 1, 4, '', head_format)
#         ws.merge_range(2, 0, 2, 1, '时间', merge_format)
#         ws.write_row(2, 2, ['课程设置', '师资方', '单位', '讲师', '备注'], merge_format)
#         startrow = 3
#         for index, course in enumerate(class_course):
#             ws.write_row(startrow + index, 0, course.get_data(), normal_format)
#
#     time_list = ['8:30-12:00', '14:00-17:30', '19:00-22:30']
#
#     for name, class_course in teacher_class.items():
#         row_datas = {}
#         for d in date_list:
#             row_datas[d] = {}
#             for t in time_list:
#                 row_datas[d][t] = ''
#
#         sheet_name = name.replace('/', '、')
#         ws = work_book.add_worksheet(sheet_name)
#         ws.default_row_height = 20
#         ws.set_column(0, 3, 22)
#         ws.set_row(0, 30)
#         ws.merge_range(0, 0, 0, 3, ('%s专家授课表' % name), head_format)
#         ws.write_row(1, 1, ['8:30-12:00', '14:00-17:30', '19:00-22:30'], merge_format)
#         for index, course in enumerate(class_course):
#             try:
#                 row_datas[course.c_date][course.c_time] += course.c_class + ' '
#             except:
#                 pass
#         startrow = 2
#         for index, d in enumerate(date_list):
#             ws.write(startrow + index, 0, d, merge_format)
#             for n in range(len(time_list)):
#                 ws.write(startrow + index, 1 + n, row_datas[d][time_list[n]], normal_format)
#
#     work_book.close()


#  在sheet表中填写时间
def write_dates_to_sheet(ws, dates, format, start_row=1):
    ws.default_row_height = 20
    ws.set_column(0, 1, 15)
    for keys, values in dates.items():
        size = len(values)
        if size > 1:
            ws.merge_range(start_row, 0, start_row + size - 1, 0, keys, format)
        else:
            ws.write(start_row, 0, keys, format)
        for value in values:
            ws.write(start_row, 1, value, format)
            start_row += 1


# 生成排课表
def export_course(teacher_dict, dates, class_dict):
    file_name = path_course_file.get() + '--排课表' + datetime.datetime.now().strftime('%H%M') + '.xlsx'
    work_book = xlsxwriter.Workbook(file_name)
    head_format = work_book.add_format(format_header)
    merge_format = work_book.add_format(format_merge)
    normal_format = work_book.add_format(format_normal)

    write_date_class_sheet(work_book, dates, class_dict)
    write_teacher_info_sheet(work_book, teacher_dict, normal_format)

    write_class_info_sheet(work_book, class_dict, teacher_dict, dates, head_format, merge_format, normal_format)
    write_teacher_time_sheet(work_book, class_dict, dates, head_format, merge_format, normal_format)

    work_book.close()
    showlog('排课表-，文件为：')
    showlog(file_name)
    pass


def get_class_course_from_schedu(book):
    class_courses = OrderedDict()
    for c in classes.get().split(' '):
        class_courses[c] = []
    if True:
        sh = book.sheet_by_name(sheet_name_class_all)
        row_one = sh.row_values(0)
        for rx in range(1, sh.nrows):
            row = sh.row(rx)
            for ix in range(2, sh.ncols):
                if row[ix].ctype == 1:
                    vaules = list(row[ix].value)
                    if len(vaules) > 0:
                        for v in vaules:
                            k = v.replace('A', '晨曦').replace('B', '晨光').replace('C', '曙光').replace('D', '朝阳').replace(
                                'E', '旭日')
                            if class_courses.__contains__(k):
                                class_courses[k].append(row_one[ix])
                elif row[ix].ctype == 2:
                    if row[ix].value == 5:
                        for c in class_courses.keys():
                            class_courses[c].append(row_one[ix])

    return class_courses


def get_class_course_from_class(book):
    class_courses = OrderedDict()
    sh = book.sheet_by_name(sheet_name_class_simple)
    for ix in range(2, sh.ncols):
        cols = sh.col_values(ix)
        if len(cols) > 1:
            class_courses[cols[0]] = cols[1:]
    return class_courses


# 生成排课简略表
def write_date_class_sheet(work_book, dates, class_courses):
    head_format = work_book.add_format(format_header)
    merge_format = work_book.add_format(format_merge)
    normal_format = work_book.add_format(format_normal)

    ws = work_book.add_worksheet(sheet_name_class_simple)
    ws.default_row_height = 20
    ws.set_column(0, 1, 15)
    ws.set_column(2, 6, 10)
    ws.set_row(0, 50)
    write_dates_to_sheet(ws, dates, merge_format)
    for index, class_name in enumerate(classes.get().split(" ")):
        ws.write(0, 2 + index, class_name, head_format)
        if class_courses.__contains__(class_name):
            ws.write_column(1, 2 + index, class_courses[class_name], normal_format)
    return ws


# 生成教师信息表
def write_teacher_info_sheet(work_book, teacher_classes, normal_format):
    ws = work_book.add_worksheet(sheet_name_teacher_info)
    ws.default_row_height = 20
    ws.set_column(0, 0, 15)
    ws.set_column(1, 1, 60)
    ws.set_column(2, 3, 15)
    start_row = 0
    for value in teacher_classes.values():
        ws.write_row(start_row, 0, value.get_data(), normal_format)
        start_row += 1


# 生成各班课程表
def write_class_info_sheet(work_book, class_dict, teacher_dict, dates, head_format, merge_format, normal_format):
    for class_name, teachers in class_dict.items():
        ws = work_book.add_worksheet(class_name)
        ws.default_row_height = 20
        ws.set_column(2, 2, 60)
        ws.set_column(3, 3, 10)
        ws.set_column(4, 4, 15)
        ws.set_column(5, 6, 10)
        ws.set_row(0, 40)
        ws.merge_range(0, 0, 0, 6, ('经营网点转型发展业务骨干培训班课程表--%s' % class_name), head_format)
        ws.merge_range(1, 0, 1, 4, '', head_format)
        ws.merge_range(2, 0, 2, 1, '时间', merge_format)
        ws.write_row(2, 2, ['课程设置', '师资方', '单位', '讲师', '备注'], merge_format)
        start_row = 3
        write_dates_to_sheet(ws, dates, merge_format, start_row)
        for index, teacher_name in enumerate(teachers):
            teacher = teacher_dict[teacher_name]
            ws.write_row(start_row + index, 2, teacher.get_class_data(), normal_format)
    pass


# 生成各专家上课时间表
def write_teacher_time_sheet(work_book, class_courses, dates, head_format, merge_format, normal_format):
    teacher_datas = {}
    for class_name, teachers in class_courses.items():
        index = 0
        for date, times in dates.items():
            for time in times:
                teacher_name = teachers[index]
                if not teacher_datas.__contains__(teacher_name):
                    teacher_datas[teacher_name] = []
                teacher_datas[teacher_name].append((date, time, class_name))
                index += 1
    time_list = ['8:00-12:00', '14:00-18:00', '19:00-23:00']
    for name, datas in teacher_datas.items():
        row_datas = OrderedDict()
        for d in dates.keys():
            row_datas[d] = OrderedDict()
            for t in time_list:
                row_datas[d][t] = ''

        sheet_name = name.replace('/', '、')
        ws = work_book.add_worksheet(sheet_name)
        ws.default_row_height = 20
        ws.set_column(0, 3, 22)
        ws.set_row(0, 30)
        ws.merge_range(0, 0, 0, 3, ('%s专家授课表' % name), head_format)
        ws.write_row(1, 1, time_list, merge_format)

        for index, data in enumerate(datas):
            if row_datas.__contains__(data[0]):
                if row_datas[data[0]].__contains__(data[1]):
                    row_datas[data[0]][data[1]] += data[2] + ' '
        start_row = 2
        index = 0
        for date, value in row_datas.items():
            ws.write(start_row + index, 0, date, merge_format)
            ws.write_row(start_row + index, 1, value.values(), normal_format)
            index += 1


def get_dates(sh):
    dates = OrderedDict()
    the_date = 'default'
    col_date = 0
    col_time = 1
    for rx in range(1, sh.nrows):
        row = sh.row(rx)
        if row[col_date].ctype == 1:
            if len(row[col_date].value) > 0:
                the_date = row[col_date].value
        elif row[col_date].ctype == 3:
            temp = xlrd.xldate.xldate_as_datetime(row[col_date].value, 0)
            the_date = temp.strftime('%m-%d')
        if not dates.__contains__(the_date):
            dates[the_date] = []
        dates[the_date].append(row[col_time].value)
    return dates


def get_teacher_info(book):
    teacher_course = {}
    sh = book.sheet_by_name(sheet_name_teacher_info)
    for rx in range(0, sh.nrows):
        row = sh.row_values(rx)
        # if len(row) == 4:
        teacher_name = row[0]
        the_teacher = Teacher(teacher_name, row[1], row[2], row[3])
        teacher_course[teacher_name] = the_teacher
    return teacher_course


def get_available(l, i, j):
    global pre_count
    if len(l[l > 9]) == 0:
        # print(l)
        pre_count += 1
        showlog('------------------')
        showlog(pre_count)
        b = l.tolist()
        for i in range(0, len(l)):
            for j in range(0, len(l[0])):
                v = b[i][j]
                if v == 0:
                    b[i][j] = ''
                elif v == 1:
                    b[i][j] = 'A'
                elif v == 2:
                    b[i][j] = 'B'
                elif v == 3:
                    b[i][j] = 'C'
                elif v == 4:
                    b[i][j] = 'D'
                elif v == 5:
                    b[i][j] = 'E'
        for i in range(0, len(l)):
            showlog('\t'.join(b[i]))
    else:
        target_x = -1
        target_y = -1
        for m in range(0, i):
            if 10 in l[m]:
                for n in range(0, j):
                    if l[m][n] == 10:
                        target_x = m
                        target_y = n
                        break
                break
        for value in range(1, 6):
            if pre_count > max_count.get():
                break
            row = l[target_x]
            col = l[:, target_y]
            if (value not in row) and (value not in col):
                l[target_x][target_y] = value
                l_temp = copy.deepcopy(l)
                get_available(l_temp, i, j)


def preview_schedu():
    global pre_count
    pre_count = 0
    if len(path_course_file.get()) <= 0:
        showinfo('提示', '先选择排课文件')
        return
    showlog(max_count.get())
    try:
        book = xlrd.open_workbook(path_course_file.get())
        sh = book.sheet_by_name(sheet_name_class_middle)
        row_count = sh.nrows
        col_count = sh.ncols
        showlog(row_count)
        showlog(col_count)
        a = np.zeros((row_count, col_count), dtype=np.int)
        for i in range(0, row_count):
            row = sh.row(i)
            for j in range(0, col_count):
                if row[j].ctype == 2:
                    if row[j].value == 1:
                        a[i][j] = 10
        get_available(a, row_count, col_count)

        showlog('已完成')
    except Exception as e:
        showlog('出错了')
        showlog(str(e))
    showlog('')


def make_course_file():
    if len(path_course_file.get()) <= 0:
        showinfo('提示', '先选择排课文件')
        return

    try:
        book = xlrd.open_workbook(path_course_file.get())
        teacher_course = get_teacher_info(book)

        if video.get() == 1:
            sh = book.sheet_by_name(sheet_name_class_all)
            class_courses = get_class_course_from_schedu(book)
        elif video.get() == 2:
            sh = book.sheet_by_name(sheet_name_class_simple)
            class_courses = get_class_course_from_class(book)
        dates = get_dates(sh)

        export_course(teacher_course, dates, class_courses)
        showlog('已完成')
    except Exception as e:
        showlog('出错了')
        showlog(str(e))
    showlog('')


def clear_text():
    logs.delete(1.0, END)


def main():
    logs.grid(row=0, column=0, rowspan=16)

    Label(root, text="分班情况\n(空格分隔)", ).grid(row=0, column=1, columnspan=1)
    Entry(root, textvariable=classes).grid(row=0, column=2, columnspan=2, )
    Separator(root, orient=HORIZONTAL).grid(row=1, column=1, columnspan=3, sticky="we")

    Button(root, text='选择课程文件', command=lambda: utils.selectFile(pathFile)).grid(row=2, column=1)
    Entry(root, textvariable=pathFile).grid(row=3, column=1)
    Button(root, text='生成排班中间表', command=make_mid_file).grid(row=4, column=1, )
    Button(root, text='清空', command=clear_text).grid(row=5, column=1, )
    Separator(root, orient=VERTICAL).grid(row=2, column=2, rowspan=3, sticky="ns")

    Button(root, text='选择排课文件', command=lambda: utils.selectFile(path_course_file)).grid(row=2, column=3)
    Entry(root, textvariable=path_course_file).grid(row=3, column=3)
    Entry(root, textvariable=max_count).grid(row=4, column=3, columnspan=1, )
    Button(root, text='预览方案', command=preview_schedu).grid(row=5, column=3, columnspan=1)
    # Radiobutton(root, text=("从'%s'预览" % sheet_name_class_middle), variable=video,value=2).grid(row=5, column=3, columnspan=1)
    Radiobutton(root, text=("从'%s'生成" % sheet_name_class_all), variable=video, value=1).grid(row=6, column=3,
                                                                                             columnspan=1)
    Radiobutton(root, text=("从'%s'生成" % sheet_name_class_simple), variable=video, value=2).grid(row=7, column=3,
                                                                                                columnspan=1)
    Button(root, text='生成课表', command=make_course_file).grid(row=8, column=3, columnspan=1)
    root.mainloop()


root = Tk()
logs = ScrolledText(root, width=40, height=40)
pathFile = StringVar()
path_course_file = StringVar()
classes = StringVar(value='晨曦 晨光 曙光 朝阳 旭日')
video = IntVar(value=1)
max_count = IntVar(value=30)
if __name__ == '__main__':
    main()
