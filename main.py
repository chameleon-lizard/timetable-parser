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

def get_day(cells, num):
    num += 2
    day_raw = [[cells[i][num + j] for i in range(len(cells))] for j in range(9)]
    day = [
            [
                [
                    str(day_raw[i][j])[str(day_raw[i][j]).index(':')+1:].replace("'", ''), 
                    str(day_raw[i+1][j])[str(day_raw[i+1][j]).index(':')+1:].replace("'", ''), 
                    str(day_raw[0][j])[str(day_raw[0][j]).index(':')+1:].replace("'", '')
                ] for j in range(len(day_raw[i]))
            ] for i in range(1, len(day_raw), 2)
        ]
    return day

printer = pp.PrettyPrinter()
book = get_book('test.xls')
second = get_sheet(book, 1)
cells = get_cells(second)
seven = get_group(cells, 1)
tuesday = get_day(cells, 1)
printer.pprint(tuesday)
