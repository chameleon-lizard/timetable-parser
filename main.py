#!/bin/python

import xlrd
import pprint as pp

def get_book(path):
    return xlrd.open_workbook(path)

def get_sheet(book, num):
    return book.sheet_by_index(num)

def get_day(sheet, num):
    # Getting raw day data
    day_dirty = [[i for i in sheet.row_values(j)] for j in range(2+num%10+num*10, 13+num%10+num*10)]

    day = [day_dirty[0]]

    for i in range(2, len(day_dirty), 2):
        # Appending time
        day.append([day_dirty[i-1][0]])
        for j in range(1, len(day_dirty[i])):
            day[i//2].append([day_dirty[i-1][j], day_dirty[i][j]])
    return day

def get_group_day(day, num):
    group = [day[0][0]]
    for i in range(1, len(day)):
        group.append([day[i][0], day[i][1 + int(num) - int(day[0][1])]])
    return group

def get_group(sheet, num):
    return [get_group_day(get_day(sheet, i), num) for i in range(5)]

printer = pp.PrettyPrinter()
book = get_book('test.xls')
sheet = get_sheet(book, 0)
day = get_day(sheet, 1)
group = get_group(sheet, 207)
printer.pprint(group)
