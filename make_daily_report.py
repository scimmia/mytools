# coding=utf-8
from datetime import date,timedelta
from collections import OrderedDict
import xlsxwriter

import datas
import utils


def main():
    header = ['日期','工作内容','备注']
    the_datas = OrderedDict()
    current_month = 1
    current_day = date(2020, 1, 1)
    while(current_day.year == 2020):
        if not the_datas.__contains__(current_month):
            the_datas[current_month] = []
        if current_day.month == current_month:
            the_datas[current_month].append(current_day.strftime('%m-%d'))
        else:
            current_month = current_day.month
        current_day += timedelta(days=1)

    workbook = xlsxwriter.Workbook('模板.xlsx')
    format_h = workbook.add_format(datas.format_header)
    for month,days in the_datas.items():
        sheet = workbook.add_worksheet('%d月' % month)
        sheet.default_row_height = 100
        sheet.default_col_width = 10
        sheet.set_row(0,20)
        sheet.set_column(1, 1, 120)
        sheet.set_column(2, 2, 30)
        sheet.write_row(0,0,header,format_h)
        sheet.write_column(1,0,days,format_h)

    workbook.close()
    print(the_datas)




if __name__ == '__main__':
    main()
