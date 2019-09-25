# coding=utf-8
import datetime
from tkinter import *
from tkinter.scrolledtext import ScrolledText
from tkinter.ttk import *
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import *

import xlrd
import xlsxwriter
from xlrd import xldate_as_tuple

orgs = {"济南": "济南", "章丘": "济南", "平阴": "济南", "济阳": "济南", "商河": "济南", "莱芜": "济南", "青岛": "青岛", "张店": "淄博", "临淄": "淄博",
        "博山": "淄博", "周村": "淄博", "桓台": "淄博", "高青": "淄博", "沂源": "淄博", "淄川": "淄博", "枣庄": "枣庄", "滕州": "枣庄", "东营": "东营",
        "垦利": "东营", "利津": "东营", "广饶": "东营", "烟台": "烟台", "龙口": "烟台", "莱阳": "烟台", "莱州": "烟台", "蓬莱": "烟台", "招远": "烟台",
        "栖霞": "烟台", "海阳": "烟台", "长岛": "烟台", "潍坊": "潍坊", "青州": "潍坊", "诸城": "潍坊", "寿光": "潍坊", "安丘": "潍坊", "高密": "潍坊",
        "昌邑": "潍坊", "昌乐": "潍坊", "临朐": "潍坊", "济宁": "济宁", "曲阜": "济宁", "兖州": "济宁", "邹城": "济宁", "汶上": "济宁", "泗水": "济宁",
        "微山": "济宁", "鱼台": "济宁", "金乡": "济宁", "嘉祥": "济宁", "梁山": "济宁", "泰山": "泰安", "岱岳": "泰安", "新泰": "泰安", "肥城": "泰安",
        "宁阳": "泰安", "东平": "泰安", "威海": "威海", "荣成": "威海", "文登": "威海", "乳山": "威海", "东港": "日照", "岚山": "日照", "莒县": "日照",
        "五莲": "日照", "滨州": "滨州", "博兴": "滨州", "邹平": "滨州", "惠民": "滨州", "阳信": "滨州", "无棣": "滨州", "德州": "德州", "乐陵": "德州",
        "禹城": "德州", "陵城": "德州", "宁津": "德州", "庆云": "德州", "临邑": "德州", "齐河": "德州", "平原": "德州", "夏津": "德州", "武城": "德州",
        "聊城": "聊城", "临清": "聊城", "高唐": "聊城", "茌平": "聊城", "东阿": "聊城", "阳谷": "聊城", "莘县": "聊城", "润昌": "聊城", "兰山": "临沂",
        "罗庄": "临沂", "河东": "临沂", "沂南": "临沂", "沂水": "临沂", "莒南": "临沂", "临沭": "临沂", "郯城": "临沂", "兰陵": "临沂", "费县": "临沂",
        "平邑": "临沂", "蒙阴": "临沂", "菏泽": "菏泽", "曹县": "菏泽", "定陶": "菏泽", "成武": "菏泽", "单县": "菏泽", "巨野": "菏泽", "郓城": "菏泽",
        "鄄城": "菏泽", "东明": "菏泽"}
cities = {"济南": {"济南": [], "章丘": [], "平阴": [], "济阳": [], "商河": [], "莱芜": []}, "青岛": {"青岛": []},
          "淄博": {"张店": [], "临淄": [], "博山": [], "周村": [], "桓台": [], "高青": [], "沂源": [], "淄川": []},
          "枣庄": {"枣庄": [], "滕州": []}, "东营": {"东营": [], "垦利": [], "利津": [], "广饶": []},
          "烟台": {"烟台": [], "龙口": [], "莱阳": [], "莱州": [], "蓬莱": [], "招远": [], "栖霞": [], "海阳": [], "长岛": []},
          "潍坊": {"潍坊": [], "青州": [], "诸城": [], "寿光": [], "安丘": [], "高密": [], "昌邑": [], "昌乐": [], "临朐": []},
          "济宁": {"济宁": [], "曲阜": [], "兖州": [], "邹城": [], "汶上": [], "泗水": [], "微山": [], "鱼台": [], "金乡": [], "嘉祥": [],
                 "梁山": []}, "泰安": {"泰山": [], "岱岳": [], "新泰": [], "肥城": [], "宁阳": [], "东平": []},
          "威海": {"威海": [], "荣成": [], "文登": [], "乳山": []}, "日照": {"东港": [], "岚山": [], "莒县": [], "五莲": []},
          "滨州": {"滨州": [], "博兴": [], "邹平": [], "惠民": [], "阳信": [], "无棣": []},
          "德州": {"德州": [], "乐陵": [], "禹城": [], "陵城": [], "宁津": [], "庆云": [], "临邑": [], "齐河": [], "平原": [], "夏津": [],
                 "武城": []}, "聊城": {"聊城": [], "临清": [], "高唐": [], "茌平": [], "东阿": [], "阳谷": [], "莘县": [], "润昌": []},
          "临沂": {"兰山": [], "罗庄": [], "河东": [], "沂南": [], "沂水": [], "莒南": [], "临沭": [], "郯城": [], "兰陵": [], "费县": [],
                 "平邑": [], "蒙阴": []},
          "菏泽": {"菏泽": [], "曹县": [], "定陶": [], "成武": [], "单县": [], "巨野": [], "郓城": [], "鄄城": [], "东明": []}}

results = {}
wrong_results = []
header_titles = []


def add_to_results(org, the_results, row):
    try:
        city = orgs.get(org)
        if not the_results.__contains__(city):
            the_results[city] = {}
        if not the_results[city].__contains__(org):
            the_results[city][org] = []
        the_results[city][org].append(row)
    except:
        pass


def init_sheet(work_book, name):
    head_format = work_book.add_format({
        'font_size': '16',
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter', })
    merge_format = work_book.add_format({
        'font_size': '12',
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter', })

    ws = work_book.add_worksheet(name)
    ws.default_row_height = 20
    ws.set_row(0, 50)
    # ws.merge_range('A1:J1', '经营网点转型发展业务骨干培训班学员名单', head_format)
    # ws.write_row(1,0,header_titles,merge_format)
    ws.set_column(len(header_titles), len(header_titles), 25)
    if name == '全部':
        ws.merge_range(0, 0, 0, len(header_titles), '经营网点转型发展业务骨干培训班学员名单', head_format)
        ws.write(1, len(header_titles) + 1, '班级', merge_format)
    else:
        ws.merge_range(0, 0, 0, len(header_titles), '经营网点转型发展业务骨干培训班%s班学员名单' % name, head_format)
    ws.write_row(1, 0, header_titles, merge_format)
    ws.write(1, len(header_titles), '签到', merge_format)
    return ws


def init_class_datas(all_datas):
    class_rows = {}
    try:
        the_class = classes.get().split(" ")
        for c in the_class:
            class_rows[c] = {}
        i = 0
        for key, values in all_datas.items():
            showlog(key)
            for k, v in values.items():
                if orgs.__contains__(k):
                    add_to_results(k, the_class[i % 5], v)
                    i = i + 1
    except:
        showlog('init_class_datas failed')
        pass
    return class_rows


def export_file(dates, teacher_classes):
    work_book = xlsxwriter.Workbook(pathFile.get() + datetime.datetime.now().strftime('%H%M') + '.xlsx')
    head_format = work_book.add_format({
        'font_size': '16',
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter', })
    merge_format = work_book.add_format({
        'font_size': '12',
        'bold': 1,
        'border': 1,
        'align': 'center',
        'valign': 'vcenter', })

    normal_format = work_book.add_format({
        'font_size': '12',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter', })

    ws = work_book.add_worksheet('课程')
    ws.default_row_height = 20
    ws.set_column(0,1,15)
    ws.set_column(2,6,10)
    ws.set_row(0, 50)
    ws.write_row(0, 2, classes.get().split(" "), head_format)
    startrow = 1
    for keys, values in dates.items():
        size = len(values)
        if size > 1:
            ws.merge_range(startrow,0,startrow+size-1,0, keys, head_format)
        else:
            ws.write(startrow, 0, keys, head_format)
        for value in values:
            # ws.write(startrow, 0, keys, head_format)
            ws.write(startrow, 1, value, merge_format)
            startrow += 1

    ws = work_book.add_worksheet('排课')
    ws.default_row_height = 20
    ws.set_column(0,1,15)
    ws.set_column(2,6,10)
    ws.write_row(0, 2, teacher_classes.keys(), normal_format)
    ws.write_row(0, 2+len(teacher_classes.keys()), classes.get().split(" "), head_format)
    startrow = 1
    for keys, values in dates.items():
        size = len(values)
        if size > 1:
            ws.merge_range(startrow,0,startrow+size-1,0, keys, head_format)
        else:
            ws.write(startrow, 0, keys, head_format)
        for value in values:
            # ws.write(startrow, 0, keys, head_format)
            ws.write(startrow, 1, value, merge_format)
            startrow += 1

    ws = work_book.add_worksheet('讲师')
    ws.default_row_height = 20
    ws.set_column(0,0,15)
    ws.set_column(1,1,60)
    ws.set_column(2,2,15)
    startrow = 0
    for keys, values in teacher_classes.items():
        ws.write(startrow, 0, keys, normal_format)
        ws.write(startrow, 1, values['course'], normal_format)
        ws.write(startrow, 2, values['org'], normal_format)
        startrow += 1

    work_book.close()


def showlog(text):
    logs.insert(END, str(text) + "\n")
    logs.update()
    print(text)


def selectFile():
    file_path = askopenfilename(filetypes=[('XLS', '*.xls;*.xlsx'), ('All Files', '*')])
    pathFile.set(file_path)

def select_course_file():
    file_path = askopenfilename(filetypes=[('XLS', '*.xls;*.xlsx'), ('All Files', '*')])
    path_course_file.set(file_path)


def preview_header():
    global startline
    try:
        book = xlrd.open_workbook(pathFile.get())
        print("The number of worksheets is {0}".format(book.nsheets))
        print("Worksheet name(s): {0}".format(book.sheet_names()))
        sh = book.sheet_by_index(0)
        print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))

        the_row = []
        for rx in range(sh.nrows):
            row = (sh.row_values(rx))
            if '姓名' in row and '手机' in row:
                showlog('found')
                the_row = row
                startline = rx + 1
                break
        listbox.delete(0, END)
        for row in the_row:
            listbox.insert(END, row)
        listbox.grid()
    except:
        showlog('打开文件失败')


def doIt():
    if len(pathFile.get()) <= 0:
        showinfo('提示', '先选择文件')
        return

    class_rows = {}
    if True:
        the_class = classes.get().split(" ")
        for c in the_class:
            class_rows[c] = {}

        book = xlrd.open_workbook(pathFile.get())
        sh = book.sheet_by_index(0)
        col_date = 0
        col_time = 1
        col_teacher = 5
        col_class = 2
        col_org = 4
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
                the_teacher = row[col_teacher].value
                if len(the_teacher) <=0:
                    the_teacher = '无'
                if teacher_classes.__contains__(the_teacher):
                    the_teacher += str(rx)
                teacher_classes[the_teacher] = {'org': row[col_org].value, 'course': row[col_class].value}

        export_file(dates, teacher_classes)
        showlog('已完成')
    # except:
    #     showlog('出错了')

    showlog('finished')


# def doIt():
#     selections = listbox.curselection()
#     if len(selections) <= 0:
#         showlog("至少选择一列")
#         return
#     global header_titles
#     header_titles = listbox.selection_get().split("\n")
#     header_titles.insert(0, '地市')
#     header_titles.insert(0, '序号')
#
#     try:
#         book = xlrd.open_workbook(pathFile.get())
#         print("The number of worksheets is {0}".format(book.nsheets))
#         print("Worksheet name(s): {0}".format(book.sheet_names()))
#         sh = book.sheet_by_index(0)
#         print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
#
#         org_index = -1
#         # for v in sh.row_values(startline):
#         for index, row in enumerate(sh.row_values(startline), 1):
#             if isinstance(row, str) and row.find('农') > 0 and row.find('商') > 0:
#                 org_index = index
#                 break
#         for rx in range(startline, sh.nrows):
#             the_row = []
#             row = (sh.row_values(rx))
#             for i in selections:
#                 the_row.append(row[i])
#             if org_index != -1:
#                 org = row[org_index].replace("农商", "").replace("商行", "").replace("银", "").replace("行", "")
#                 if orgs.__contains__(org):
#                     add_to_results(org, results, the_row)
#                 else:
#                     wrong_results.append(the_row)
#             else:
#                 wrong_results.append(the_row)
#     except:
#         showlog('出错了')
#     class_rows = init_class_datas(results)
#     export_file(results, class_rows)
#     showlog('已完成')
#
#     showlog('finished')


def startIt():
    if len(pathFile.get()) <= 0:
        showinfo('提示', '先选择文件')
        return
    doIt()


def wirte_sheet(the_sheet, normal_format, all_rows):
    startrow = 2
    for city in cities.keys():
        if all_rows.__contains__(city):
            the_city = all_rows[city]
            for org in cities[city].keys():
                if the_city.__contains__(org):
                    for temp in the_city[org]:
                        the_sheet.write(startrow, 0, str(startrow - 1), normal_format)
                        the_sheet.write(startrow, 1, city, normal_format)
                        the_sheet.write_row(startrow, 2, temp, normal_format)
                        startrow = startrow + 1
    return startrow


def wirte_summary(sheet_summary, normal_format, all_rows):
    m = 0
    for city in cities.keys():
        if all_rows.__contains__(city):
            the_city = all_rows[city]
            size = len(the_city)
            if size > 1:
                a = 'A%d:A%d' % (m + 1, m + size)
                b = 'B%d:B%d' % (m + 1, m + size)
                sheet_summary.merge_range(a, city, normal_format)
                sheet_summary.merge_range(b, '=SUM(D%d:D%d)' % (m + 1, m + size), normal_format)
            else:
                sheet_summary.write(m, 0, city, normal_format)
                sheet_summary.write(m, 1, '=SUM(D%d:D%d)' % (m + 1, m + size), normal_format)
            for org in cities[city].keys():
                if the_city.__contains__(org):
                    sheet_summary.write(m, 2, org, normal_format)
                    sheet_summary.write(m, 3, len(the_city[org]), normal_format)
                    m = m + 1
    sheet_summary.write(m, 0, '合计', normal_format)
    sheet_summary.write(m, 1, '=SUM(D:D)', normal_format)


def make_all_from_class_rows(class_rows):
    all_rows = {}
    for k, v in class_rows.items():
        for ks, vs in v.items():  # key city
            if not all_rows.__contains__(ks):
                all_rows[ks] = {}
            for kss, vss in vs.items():  # key org
                if not all_rows[ks].__contains__(kss):
                    all_rows[ks][kss] = []
                all_rows[ks][kss] += vss
    return all_rows


def write_to_file(all_rows, class_rows, wrong_results):
    work_book = xlsxwriter.Workbook(pathFile.get() + '--' + datetime.datetime.now().strftime('%H%M') + '.xlsx')
    normal_format = work_book.add_format({
        'font_size': '12',
        'border': 1,
        'align': 'center',
        'valign': 'vcenter', })
    worksheets = {}
    ws_all = init_sheet(work_book, '全部')
    worksheets['全部'] = ws_all
    startrow = wirte_sheet(ws_all, normal_format, all_rows)
    for wrong_result in wrong_results:
        ws_all.write_row(startrow, 0, wrong_result, normal_format)
        ws_all.write(startrow, len(wrong_result), '错误', normal_format)
        startrow = startrow + 1

    the_class = classes.get().split(" ")
    for c in the_class:
        sheet_temp = init_sheet(work_book, c)
        worksheets[c] = sheet_temp
        wirte_sheet(sheet_temp, normal_format, class_rows[c])

    sheet_city = work_book.add_worksheet('按地市汇总')
    worksheets['按地市汇总'] = sheet_city
    wirte_summary(sheet_city, normal_format, all_rows)

    work_book.close()
    showlog('over')


def make_course():
    if len(path_course_file.get()) <= 0:
        showinfo('提示', '先选择文件')
        return
    class_rows = {}
    if True:
        the_class = classes.get().split(" ")
        for c in the_class:
            class_rows[c] = []

def main():
    logs.grid(row=0, column=0, rowspan=6)

    Label(root, text="分班情况\n(空格分隔)", ).grid(row=0, column=1, columnspan=1)
    Entry(root, textvariable=classes).grid(row=0, column=2 ,columnspan=2,)
    Separator(root, orient=HORIZONTAL).grid(row=1, column=1, columnspan=3, sticky="we")


    Button(root, text='选择课程文件', command=selectFile).grid(row=2, column=1)
    Entry(root, textvariable=pathFile).grid(row=3, column=1)
    Button(root, text='生成排班', command=doIt).grid(row=4, column=1, )
    Separator(root, orient=VERTICAL).grid(row=2, column=2, rowspan=3, sticky="ns")

    Button(root, text='选择排课文件', command=select_course_file()).grid(row=2, column=3)
    Entry(root, textvariable=path_course_file).grid(row=3, column=3)
    Button(root, text='生成课表', command=make_course).grid(row=4, column=3, columnspan=1)
    root.mainloop()


root = Tk()
logs = ScrolledText(root, width=100, height=40)
pathFile = StringVar()
path_course_file = StringVar()
classes = StringVar()
classes.set('晨曦 晨光 曙光 朝阳 旭日')
chVarDis = BooleanVar()
check1 = Checkbutton(root, text="部分修改", variable=chVarDis)
listbox = Listbox(root, selectmode=MULTIPLE)
startline = -1
if __name__ == '__main__':
    main()
