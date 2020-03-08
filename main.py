#!/bin/python

import xlrd
import pprint as pp

def get_book(path):
    return xlrd.open_workbook(path)

def get_sheet(book, num):
    return book.sheet_by_index(num)

def get_cells(sheet):
    return [[sheet.cell(j, i) for j in range(60)] for i in range(7)]

def get_group(cells, num):
    return cells[num]

def get_day(sheet, num):
    day_dirty = [[i for i in sheet.row_values(j)] for j in range(2+num*9, 11+num*9)]
    day = [day_dirty[0]]
    for i in range(1, len(day_dirty) - 1, 2):
        day.append(day_dirty[i][0])
        if i == 0:
            day.append(day_dirty[i])
            continue
        else:
            for j in range(1, len(day_dirty[i])):
                day.append([day_dirty[i][j], day_dirty[i+1][j]])

    return day

printer = pp.PrettyPrinter()
book = get_book('test.xls')
second = get_sheet(book, 1)
cells = get_cells(second)
seven = get_group(cells, 1)
day = get_day(second, 1)
printer.pprint(day)
