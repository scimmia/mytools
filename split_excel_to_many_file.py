# coding=utf-8
import datetime
from tkinter import *
from tkinter.scrolledtext import ScrolledText
from tkinter.ttk import *
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import *

import xlrd
import xlsxwriter

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


def get_sheet(work_book, name):
    worksheet = work_book.get_worksheet_by_name(name)
    if worksheet is None:
        worksheet = work_book.add_worksheet(name)
        worksheet.default_row_height = 20
    return worksheet




def showlog(text):
    logs.insert(END, str(text) + "\n")
    logs.update()
    print(text)


def selectFile():
    file_path = askopenfilename(filetypes=[('XLS', '*.xls;*.xlsx'), ('All Files', '*')])
    pathFile.set(file_path)


def doIt():
    name_org = {}
    try:
        work_book = xlsxwriter.Workbook(pathFile.get() + '--' + datetime.datetime.now().strftime('%H%M') + '.xlsx')
        normal_format = work_book.add_format({
            'font_size': '12',
            'border': 1,
            'align': 'center',
            'valign': 'vcenter', })


        book = xlrd.open_workbook(pathFile.get())
        sh = book.sheet_by_index(0)
        current_index = 0
        current_sheet = None
        for rx in range(sh.nrows):
            row = (sh.row_values(rx))
            name = row[1]
            if len(name) > 0:
                current_sheet = get_sheet(work_book,name)
                current_index = 0
            current_sheet.write_row(current_index,0,row,normal_format)
            current_index += 1
        work_book.close()

    except:
        showlog('打开文件失败')
    showlog('finished')


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
    m = 1
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

def wirte_summary_all(sheet_summary, normal_format, all_rows):
    m = 0
    for city in cities.keys():
        the_city = cities[city]
        size = len(the_city)
        if size > 1:
            a = 'A%d:A%d' % (m + 1, m + size)
            b = 'B%d:B%d' % (m + 1, m + size)
            sheet_summary.merge_range(a, city, normal_format)
            sheet_summary.merge_range(b, '=SUM(D%d:D%d)' % (m + 1, m + size), normal_format)
        else:
            sheet_summary.write(m, 0, city, normal_format)
            sheet_summary.write(m, 1, '=SUM(D%d:D%d)' % (m + 1, m + len(the_city)), normal_format)
        for org in cities[city].keys():
            sheet_summary.write(m, 2, org, normal_format)
            if all_rows.__contains__(city) and all_rows[city].__contains__(org):
                sheet_summary.write(m, 3, len(all_rows[city][org]), normal_format)
            else:
                sheet_summary.write(m, 3, 0, normal_format)
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


def main():
    logs.grid(row=0, column=0, rowspan=6)

    Button(root, text='选择报名文件', command=selectFile).grid(row=0, column=1)
    Entry(root, textvariable=pathFile).grid(row=0, column=2)
    Button(root, text='开始', command=startIt).grid(row=2, column=1, columnspan=1)
    root.mainloop()


root = Tk()
logs = ScrolledText(root, width=40, height=30)
pathFile = StringVar()
startline = -1
if __name__ == '__main__':
    main()
