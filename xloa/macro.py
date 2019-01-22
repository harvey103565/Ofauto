#! /usr/bin/python3
# -*- coding: utf-8 -*-


class Macro(object):
    """
    Object presents a macro(or user defined function) in workbook. Excel function will be involved as a result of
    calling this object.
    """
    def __init__(self, app, macro):
        self._app = app
        self.macro = macro

    def run(self, *args):
        return self._app.Api.Run(self.macro, *args)

    __call__ = run

