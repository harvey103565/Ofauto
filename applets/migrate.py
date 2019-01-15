#! /usr/bin/python3
# -*- coding: utf-8 -*-


import json
import base64

from sys import argv

from timber import timber

from applets import XlController
from applets import AppError


_FILLED_IN_ = 'Filled-in'
_OVERWRITTEN_ = 'Overwritten'
_DIFFERENT_ = 'Different'
_MISMATCHED_ = 'Mismatched'
_UNUSED_ = 'Unused'
_IGNORED_ = 'Ignored'
_REDUNDANT_ = 'Redundant'

# _UNUSED_ MUST BE the last one in tuple. It is updated at last, in a loop follow the tuple's order,
#  after all other terms were counted.
data_sequence = (_FILLED_IN_, _OVERWRITTEN_, _DIFFERENT_, _MISMATCHED_, _IGNORED_, _REDUNDANT_, _UNUSED_)

color_names = ('blue', 'green', 'wheat', 'tomato', 'sky', 'purple', 'navy')

color_values = ((79, 129, 189),         # blue
                (155, 87, 189),         # green
                (247, 150, 70),         # wheat
                (192, 80, 77),          # tomato
                (75, 172, 198),         # sky
                (128, 100, 162),        # purple
                (31, 73, 125))          # navy


color_interpreter = dict(zip(color_names, color_values))


@XlController(*argv)
class XlMigrator(object):
    def __init__(self, app, sheet):
        self.summary = dict(zip(data_sequence, (0, 0, 0, 0, 0, 0, 0)))
        self.color_schema = dict(zip(data_sequence, color_names))

        self.xl_app = app
        self.xl_sheet = sheet

        self.parstr = base64.b64decode(argv[6])

        param = json.loads(self.parstr)

        self.over_writing = param['Overwriting']
        self.target = param['To']
        self.sources = param['From']

        self.xl_books = dict()
        self.xl_sheets = dict()

        self.data = dict()
        self.redundants = set()

    def __call__(self, *args, **kwargs):
        for instruction in self.sources:
            xls_src = self.GetWorksheet(instruction['book'], instruction['sheet'])

            self.LoadSourceValues(xls_src, instruction['key'], instruction['value'])

        xls_tar = self.GetWorksheet(self.target['book'], self.target['sheet'])
        self.WriteTargetValues(xls_tar, self.target['key'], self.target['value'])

        summary = self.MakeSummary()
        return summary

    @staticmethod
    def WorkingCells(sht, key_addr, value_addr):
        cell_key = sht.Range(key_addr)
        cell_value = sht.Range(value_addr)

        row_key_value = [c.Row for r in (cell_key, cell_value) for c in r]
        if len(set(row_key_value)) > 1:
            raise AppError(r'<KEY> and <VALUE> keywords are not in the same row.')

        column_key = sorted([c.Column for c in cell_key])
        column_value = sorted([c.Column for c in cell_value])

        for row in range(row_key_value[0] + 1, sht.UsedRange.Rows.Count + 1):
            yield ((row, column_key[0]), (row, column_key[-1])), ((row, column_value[0]), (row, column_value[-1]))

    def LoadSourceValues(self, sheet, key_addr, value_addr):
        for r in XlMigrator.WorkingCells(sheet, key_addr, value_addr):
            key, value = XlMigrator.RowValues(sheet, *r[0]), XlMigrator.RowValues(sheet, *r[1])

            if not any(key) or not any(value):
                continue

            if key in self.data:
                self.redundants.add(key)
                timber.info('Multiple value {0}: {1} vs {2}'.format(key, value, self.data[key]))
                continue

            self.data[key] = value
            timber.info('Find {0}: {1}'.format(key, value))

    def WriteTargetValues(self, sheet, key_addr, value_addr):
        for r in XlMigrator.WorkingCells(sheet, key_addr, value_addr):
            key, value = XlMigrator.RowValues(sheet, *r[0]), XlMigrator.RowValues(sheet, *r[1])

            if not any(key) or key not in self.data:
                self.UpdateCellRecord(_MISMATCHED_, r[0])
                continue

            if key in self.redundants:
                self.UpdateCellRecord(_REDUNDANT_, r[0])

            data = self.data[key]

            for i in range(len(value)):
                timber.info('Compare: {0}: <value>{1} vs <new value>{2}, '.format(key, value[i], data[i]))

                if not value[i]:
                    self.UpdateCellRecord(_FILLED_IN_, r[1][i + 1], data[i])
                    continue

                if data[i] == value[i]:
                    self.UpdateCellRecord(_IGNORED_, r[1][i + 1])
                    continue

                if self.over_writing:
                    self.UpdateCellRecord(_OVERWRITTEN_, r[1][i + 1], data[i])
                else:
                    self.UpdateCellRecord(_DIFFERENT_, r[1][i + 1])

    def UpdateCellRecord(self, term, row, data=None):
        self.summary[term] += 1

        if not data:
            value = XlMigrator.RowValues(row)
            if not any(value):
                return
        else:
            row.Value = data

        color_name = self.color_schema[term]
        color_value = color_interpreter[color_name]
        row.Color = color_value

    def GetWorksheet(self, book_name, sheet_name):
        book_sheet = (book_name, sheet_name)
        if book_sheet in self.xl_sheets:
            return self.xl_sheets[book_sheet]

        if book_name in self.xl_books:
            xls_book = self.xl_books[book_name]
        else:
            xls_book = self.xl_app.Workbooks[book_name]
            self.xl_books[book_name] = xls_book

        xls_sheet = xls_book.Worksheets[sheet_name]
        self.xl_sheets[book_sheet] = xls_sheet

        return xls_sheet

    def MakeSummary(self):
        total = len(self.data)
        unused = total

        mis = self.summary[_MISMATCHED_]
        rdd = self.summary[_REDUNDANT_]

        num = -(rdd + mis)
        for k in data_sequence:
            unused -= num
            num = self.summary[k]

        self.summary[_UNUSED_] = unused

        summary = (total + mis, mis, total,
                   self.summary[_FILLED_IN_],
                   self.summary[_OVERWRITTEN_],
                   self.summary[_DIFFERENT_],
                   self.summary[_IGNORED_],
                   unused,
                   rdd)

        timber.info(
            """
            
            ------------------------------------------
            运行结果：
                共计处理数据{0}条，
                    目标表格中{1}条未能匹配源数据
                    从源数据发现数据{2}条
                        填入数据{3}条
                        改写数据{4}条
                        保留有差别的数据{5}条
                        跳过不变的数据{6}条
                        未使用的数据{7}条
                        ---
                        另有{8}条数据存在多值
            ------------------------------------------
            """.format(*summary))

        return summary

    @staticmethod
    def RowValues(sheet, *coords) -> tuple:
        cells = sheet(*coords)
        return tuple(str(v).strip() if v else '' for v in cells.Value[0])


timber.basicConfig(level=timber.INFO,
                   format='%(asctime)s %(filename)s[line:%(lineno)d] %(levelname)s %(message)s',
                   datefmt='%a, %d %b %Y %H:%M:%S')

migrator = XlMigrator()
migrator()
