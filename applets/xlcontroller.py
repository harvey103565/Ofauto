#! /usr/bin/python3
# -*- coding: utf-8 -*-

import json

from pythoncom import com_error
from timber import timber

from os import path
from sys import exit
from sys import modules

from xloa import App
from xloa import XlError


def XlController(*args):

    def XlDecorator(cls):
        try:
            app = App()
            xl_book = app.Workbooks[path.basename(args[1])]
            xl_sheet = xl_book.Worksheets[args[5]]
        except com_error as e:
            timber.critical('无法建立到Excel的联接，可能是由于Excel处于锁定，或系统占用状态，需重启电脑后重试。')
        except XlError as e:
            timber.exception(e)

        def __init__(XlHander):
            cls.__init__(XlHander, app, xl_sheet)

        def __call__(XlHandler, *args, **kwargs):
            try:
                result_call_back = app.Macro('Migration.OnResultCallBack')
                for result in cls.__call__(XlHandler, *args, **kwargs):
                    result_call_back(0, json.dumps(result))
                ret_value = 0
            except com_error as exp:
                timber.exception(exp)
                ret_value = 1
            except Exception as exp:
                timber.exception(exp)
                ret_value = 2
            finally:
                mod_path = modules[cls.__module__].__file__
                mod_dir = path.dirname(mod_path)
                mod_name = path.splitext(path.basename(mod_path))[0]

                exit_call_back = app.Macro('Migration.OnExitCallBack')
                exit_call_back(1, mod_dir, mod_name)

            exit(ret_value)

        return type('XlHandler', (cls,), {'__init__': __init__, '__call__': __call__})

    return XlDecorator
