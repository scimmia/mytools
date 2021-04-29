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


def check_value(l, i, j, value):
    row = l[i]
    col = l[:, j]
    return (value not in row) and (value not in col)


osos = 0


def get_available(l, i, j):
    global osos

    if len(l[l > 9]) == 0:
        # print(l)
        osos += 1
        print(osos)
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
            print('\t'.join(b[i]))
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
            if osos > 10:
                break
            if check_value(l, target_x, target_y, value):
                l[target_x][target_y] = value
                l_temp = copy.deepcopy(l)
                get_available(l_temp, i, j)

    # result = -1
    # row = l[i]
    # col = l[:, j]
    # for value in range(6, 11):
    #     if (value not in row) and (value not in col):
    #         result = value
    #
    # return result


def do(m, n):
    b = [[10, 10, 10, 10, 10, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
         [10, 10, 10, 10, 10, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
         [10, 10, 0, 10, 10, 10, 0, 0, 0, 0, 0, 0, 0, 0, 0],
         [10, 10, 10, 10, 0, 0, 0, 0, 10, 0, 0, 0, 0, 0, 0],
         [0, 10, 10, 0, 10, 0, 10, 0, 10, 0, 0, 0, 0, 0, 0],
         [10, 0, 0, 10, 0, 0, 10, 10, 10, 0, 0, 0, 0, 0, 0],
         [0, 0, 10, 0, 10, 10, 0, 10, 0, 10, 0, 0, 0, 0, 0],
         [0, 0, 0, 0, 0, 10, 10, 10, 10, 10, 0, 0, 0, 0, 0],
         [0, 0, 0, 0, 0, 10, 10, 10, 10, 10, 0, 0, 0, 0, 0],
         [0, 0, 0, 0, 0, 10, 10, 10, 0, 10, 10, 0, 0, 0, 0],
         [0, 0, 0, 0, 0, 0, 0, 0, 0, 10, 0, 10, 10, 10, 10],
         [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 10, 10, 10, 10, 10],
         [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 10, 10, 10, 10, 10],
         [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 10, 10, 10, 10, 10],
         [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 10, 10, 10, 10, 10]]
    m = len(b)
    n = len(b[0])

    a = np.zeros((m, n), dtype=np.int)

    for i in range(0, m):
        for j in range(0, n):
            a[i][j] = b[i][j]
    # a = np.random.randint(0, 2, size=[m, n])
    # a = a * 10
    # b = np.transpose(a)

    # a[0] = [0, 0, 10, 10, 10, 0, 10, 10]
    # a[1] = [10, 10, 10, 10, 10, 0, 0, 0]
    # a[2] = [10, 10, 10, 10, 10, 0, 0, 0]
    # a[3] = [10, 10, 10, 10, 10, 0, 0, 0]
    # a[4] = [10, 10, 10, 10, 10, 0, 0, 0]
    print(a)
    get_available(a, m, n)


    # get_available(a, 15, 15)

    # # print(b)
    # for i in range(0, m):
    #     for j in range(0, n):
    #         a[i][j] = i % 6 + j % 5
    #         # b[j][i] = i % 6 + j % 5
    # print(a)
    # print(b)
    # b = np.transpose(a)
    # b[3][3] = 234
    # print(a)
    # print(b)
    print(a[2:4, 3])
    print(a[4, 1:4])
    print(len(a[a > 10]))


def main():
    do(5, 8)
    # print(eye(4))


if __name__ == '__main__':
    main()
