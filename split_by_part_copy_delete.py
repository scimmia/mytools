# coding=utf-8
# 按地市法人分割汇总文件
import datetime
import os
from tkinter import *
from tkinter.scrolledtext import ScrolledText
from tkinter.ttk import *
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import *
from utils import get_city_org,orgs,cities
import xlrd
import xlsxwriter
from shutil import copyfile
import openpyxl
import _thread
# orgs = {"济南": "济南", "章丘": "济南", "平阴": "济南", "济阳": "济南", "商河": "济南", "莱芜": "济南", "青岛": "青岛", "张店": "淄博", "临淄": "淄博",
#         "博山": "淄博", "周村": "淄博", "桓台": "淄博", "高青": "淄博", "沂源": "淄博", "淄川": "淄博", "枣庄": "枣庄", "滕州": "枣庄", "东营": "东营",
#         "垦利": "东营", "利津": "东营", "广饶": "东营", "烟台": "烟台", "龙口": "烟台", "莱阳": "烟台", "莱州": "烟台", "蓬莱": "烟台", "招远": "烟台",
#         "栖霞": "烟台", "海阳": "烟台", "长岛": "烟台", "潍坊": "潍坊", "青州": "潍坊", "诸城": "潍坊", "寿光": "潍坊", "安丘": "潍坊", "高密": "潍坊",
#         "昌邑": "潍坊", "昌乐": "潍坊", "临朐": "潍坊", "济宁": "济宁", "曲阜": "济宁", "兖州": "济宁", "邹城": "济宁", "汶上": "济宁", "泗水": "济宁",
#         "微山": "济宁", "鱼台": "济宁", "金乡": "济宁", "嘉祥": "济宁", "梁山": "济宁", "泰山": "泰安", "岱岳": "泰安", "新泰": "泰安", "肥城": "泰安",
#         "宁阳": "泰安", "东平": "泰安", "威海": "威海", "荣成": "威海", "文登": "威海", "乳山": "威海", "东港": "日照", "岚山": "日照", "莒县": "日照",
#         "五莲": "日照", "滨州": "滨州", "博兴": "滨州", "邹平": "滨州", "惠民": "滨州", "阳信": "滨州", "无棣": "滨州", "德州": "德州", "乐陵": "德州",
#         "禹城": "德州", "陵城": "德州", "宁津": "德州", "庆云": "德州", "临邑": "德州", "齐河": "德州", "平原": "德州", "夏津": "德州", "武城": "德州",
#         "聊城": "聊城", "临清": "聊城", "高唐": "聊城", "茌平": "聊城", "东阿": "聊城", "阳谷": "聊城", "莘县": "聊城", "润昌": "聊城", "兰山": "临沂",
#         "罗庄": "临沂", "河东": "临沂", "沂南": "临沂", "沂水": "临沂", "莒南": "临沂", "临沭": "临沂", "郯城": "临沂", "兰陵": "临沂", "费县": "临沂",
#         "平邑": "临沂", "蒙阴": "临沂", "菏泽": "菏泽", "曹县": "菏泽", "定陶": "菏泽", "成武": "菏泽", "单县": "菏泽", "巨野": "菏泽", "郓城": "菏泽",
#         "鄄城": "菏泽", "东明": "菏泽"}
# cities = {"济南": {"济南": [], "章丘": [], "平阴": [], "济阳": [], "商河": [], "莱芜": []}, "青岛": {"青岛": []},
#           "淄博": {"张店": [], "临淄": [], "博山": [], "周村": [], "桓台": [], "高青": [], "沂源": [], "淄川": []},
#           "枣庄": {"枣庄": [], "滕州": []}, "东营": {"东营": [], "垦利": [], "利津": [], "广饶": []},
#           "烟台": {"烟台": [], "龙口": [], "莱阳": [], "莱州": [], "蓬莱": [], "招远": [], "栖霞": [], "海阳": [], "长岛": []},
#           "潍坊": {"潍坊": [], "青州": [], "诸城": [], "寿光": [], "安丘": [], "高密": [], "昌邑": [], "昌乐": [], "临朐": []},
#           "济宁": {"济宁": [], "曲阜": [], "兖州": [], "邹城": [], "汶上": [], "泗水": [], "微山": [], "鱼台": [], "金乡": [], "嘉祥": [],
#                  "梁山": []}, "泰安": {"泰山": [], "岱岳": [], "新泰": [], "肥城": [], "宁阳": [], "东平": []},
#           "威海": {"威海": [], "荣成": [], "文登": [], "乳山": []}, "日照": {"东港": [], "岚山": [], "莒县": [], "五莲": []},
#           "滨州": {"滨州": [], "博兴": [], "邹平": [], "惠民": [], "阳信": [], "无棣": []},
#           "德州": {"德州": [], "乐陵": [], "禹城": [], "陵城": [], "宁津": [], "庆云": [], "临邑": [], "齐河": [], "平原": [], "夏津": [],
#                  "武城": []}, "聊城": {"聊城": [], "临清": [], "高唐": [], "茌平": [], "东阿": [], "阳谷": [], "莘县": [], "润昌": []},
#           "临沂": {"兰山": [], "罗庄": [], "河东": [], "沂南": [], "沂水": [], "莒南": [], "临沭": [], "郯城": [], "兰陵": [], "费县": [],
#                  "平邑": [], "蒙阴": []},
#           "菏泽": {"菏泽": [], "曹县": [], "定陶": [], "成武": [], "单县": [], "巨野": [], "郓城": [], "鄄城": [], "东明": []}}


def selectFile():
    file_path = askopenfilename(filetypes=[('XLS', '*.xls;*.xlsx'), ('All Files', '*')])
    pathFile.set(file_path)


def showlog(text):
    logs.insert(END, str(text) + "\n")
    logs.update()
    print(text)


def add_to_results(the_results, org, row):
    if not the_results.__contains__(org):
        the_results[org] = []
    the_results[org].append(row)


def write_to_file(results):
    name, suffix = os.path.splitext(pathFile.get())
    file_folder = name + '--' + datetime.datetime.now().strftime('%H%M')
    # file_folder = pathFile.get() + '--' + datetime.datetime.now().strftime('%H%M')
    if not os.path.exists(file_folder):
        os.makedirs(file_folder)
    for city, datas in results.items():
        # filename = file_folder + os.sep + city + '.xlsx'
        if video_split_type.get() == 1:
            filename = os.sep.join([file_folder, city + '.xlsx'])
            work_book = xlsxwriter.Workbook(filename)
            normal_format = work_book.add_format({
                'font_size': '12',
                'border': 1,
                'align': 'center',
                'valign': 'vcenter', })
            sheet_city = work_book.add_worksheet('sheet1')
            m = 1
            for org, rows in datas.items():
                for row in rows:
                    sheet_city.write_row(m, 0, row, normal_format)
                    m = m + 1
            work_book.close()
            showlog(filename)
        elif video_split_type.get() == 2:
            file_folder_city = os.sep.join([file_folder, city])
            # file_folder = pathFile.get() + '--' + datetime.datetime.now().strftime('%H%M')
            if not os.path.exists(file_folder_city):
                os.makedirs(file_folder_city)
            for org, rows in datas.items():
                filename = os.sep.join([file_folder_city, org + '.xlsx'])
                work_book = xlsxwriter.Workbook(filename)
                normal_format = work_book.add_format({
                    'font_size': '12',
                    'border': 1,
                    'align': 'center',
                    'valign': 'vcenter', })
                sheet_city = work_book.add_worksheet('sheet1')
                m = 1
                for row in rows:
                    sheet_city.write_row(m, 0, row, normal_format)
                    m = m + 1
                work_book.close()
                showlog(filename)

    showlog('已完成')

def get_city_org(name):
    if not isinstance(name, str):
        return None
    else:
        all_orgs = orgs.keys()
        if name[4:6] in all_orgs:
            org = name[4:6]
        elif name[2:4] in all_orgs:
            org = name[2:4]
        elif name[0:2] in all_orgs:
            org = name[0:2]
        else:
            return None
        return [orgs[org], org]

def startIt():
    row_index = start_line.get() - 1
    org_index = col_line.get() - 1
    if len(pathFile.get()) <= 0:
        showinfo('提示', '先选择文件')
        return
    results = {}
    try:
        book = xlrd.open_workbook(pathFile.get())
        sh = book.sheet_by_index(0)
        start = row_index
        end = sh.nrows
        for rx in range(row_index, sh.nrows):
            row = sh.row_values(rx)
            name = row[org_index]
            if isinstance(name, str):
                org = row[org_index].replace("农商", "").replace("商行", "").replace("银", "").replace("行", "")
                add_to_results(results, org, rx)
        name, suffix = os.path.splitext(pathFile.get())
        file_folder = name + '--' + datetime.datetime.now().strftime('%H%M')
        # file_folder = pathFile.get() + '--' + datetime.datetime.now().strftime('%H%M')
        if not os.path.exists(file_folder):
            os.makedirs(file_folder)
        for city, lines in results.items():
            filename = os.sep.join([file_folder, city + suffix])
            copyfile(pathFile.get(), filename)
            try:
                book = openpyxl.load_workbook(filename)
                sheet = book.get_active_sheet()
                for rx in range(end,start,-1):
                    if rx not in lines:
                        sheet.delete_rows(rx)
                book.save(filename)
                showlog(city+'已完成')
            except Exception as e:
                showlog(str(e))
            pass
        showlog('已完成')
    except Exception as e:
        showlog('出错了')
        showlog(str(e))
    return


def begin_work():
    try:
        _thread.start_new_thread(startIt, ())
    except:
        print("Error: 无法启动线程")


def main():
    logs.grid(row=0, column=0, rowspan=6)

    Button(root, text='选择文件', command=selectFile).grid(row=0, column=1)
    Entry(root, textvariable=pathFile).grid(row=0, column=2)
    Label(root, text='开始行号', ).grid(row=1, column=1)
    Entry(root, textvariable=start_line).grid(row=1, column=2)
    Label(root, text='哪一列', ).grid(row=2, column=1)
    Entry(root, textvariable=col_line).grid(row=2, column=2)
    Radiobutton(root, text=('按地市划分'), variable=video_split_type, value=1).grid(row=3, column=1)
    Radiobutton(root, text=('按法人划分'), variable=video_split_type, value=2).grid(row=3, column=2)
    # Radiobutton(root, text=('分割成文件'), variable=video_file_type, value=1).grid(row=4, column=1)
    # Radiobutton(root, text=('分割成sheet'), variable=video_file_type, value=2).grid(row=4, column=2)

    Button(root, text='开始', command=begin_work).grid(row=5, column=1, columnspan=1)
    root.mainloop()


root = Tk()
logs = ScrolledText(root, width=40, height=30)
pathFile = StringVar()
video_split_type = IntVar(value=1)
video_file_type = IntVar(value=1)
start_line = IntVar(value=1)
col_line = IntVar(value=1)
if __name__ == '__main__':
    main()
