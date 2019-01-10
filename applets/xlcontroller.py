#! /usr/bin/python3
# -*- coding: utf-8 -*-

import json

from pythoncom import com_error
from timber import timber

from os import path
from sys import exit

from xloa import App


def XlController(*args):

    def XlDecorator(cls):
        app = App()
        xl_book = app.Workbooks[path.basename(args[1])]
        xl_sheet = xl_book.Worksheets[args[5]]

        def __init__(XlHander):
            timber.info('XlHandler.__init__()')
            cls.__init__(XlHander, app, xl_sheet)

        def __call__(XlHandler, *args, **kwargs):
            timber.info('XlHandler.__call__({})'.format(repr(cls)))

            try:
                result = cls.__call__(XlHandler, *args, **kwargs)

                result_call_back = app.Macro('Migration.OnResultCallBack')
                ret_value = result_call_back(0, json.dumps(result))

                XlHandler.MakeSummary()
            except com_error as exp:
                timber.exception(exp)
                ret_value = 1
            except Exception as exp:
                timber.exception(exp)
                ret_value = 2
            finally:
                exit_call_back = app.Macro('Migration.OnExitCallBack')
                exit_call_back()

            exit(ret_value)

        return type('XlHandler', (cls,), {'__init__': __init__, '__call__': __call__})

    return XlDecorator
