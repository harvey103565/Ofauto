#! /usr/bin/python3
# -*- coding: utf-8 -*-

import xlwings as xw

from os import path

from xlwings._xlwindows import App
from xlwings._xlwindows import COMRetryObjectWrapper

from win32com.client import Dispatch


class Book(object):
    def __init__(self, xls):
        self.book = Book.try_app(xls)

    def __contains__(self, module: str):
        return module in [sht.name for sht in self.book.sheets]

    def __getitem__(self, item):
        return self.book.sheets[item]

    def __iter__(self):
        yield from self.book.sheets

    @staticmethod
    def load_mapping(sheet, expr, separator: str):
        index = dict(p.split(separator) for p in expr.split(','))
        return dict(((sheet.range(k).value, sheet.range(v).value) for (k, v) in index.items()))

    @staticmethod
    def dump_column(sheet, column: iter, address: str):
        used = Book.used_range(sheet)

        index = used[0]
        sheet.clear()
        cells = list([index])

        for cell in column:
            row = [''] * len(index)
            row[0] = cell
            cells.append(row)

        sheet.range(address).value = cells

    @staticmethod
    def used_range(sheet) -> tuple:
        return sheet.impl.xl.usedRange()

    @staticmethod
    def try_ole(cls_name: str, xls: str) -> xw.Book:
        app = xw.App(impl=App(
            xl=COMRetryObjectWrapper(Dispatch(cls_name))))
        return app.books(xls)

    @staticmethod
    def try_app(fn: str) -> xw.Book:
        xls = path.basename(fn)
        try:
            # Called from WPS later versions, for earlier version, it is 'et.Application'
            return Book.try_ole('Ket.Application', xls)
        except KeyError:
            pass

        try:
            # Called from Excel
            return Book.try_ole('Excel.Application', xls)
        except KeyError as e:
            raise OSError(e)
