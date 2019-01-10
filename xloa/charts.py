#! /usr/bin/python3
# -*- coding: utf-8 -*-


class Chart(object):
    """
    Class represents a chart in Workbook or specified worksheet
    """
    def __init__(self, impl):
        self._impl = impl

    @property
    def Api(self):
        """
        Return com object
        :return: com object
        """
        return self._impl


class Charts(object):
    """
    Class represents a collection of all charts in Workbook or specified worksheet
    """
    def __init__(self, impl):
        self._impl = impl

    @property
    def Api(self):
        """
        Return com object
        :return: com object
        """
        return self._impl
