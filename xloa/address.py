#! /usr/bin/python3
# -*- coding: utf-8 -*-


class Address(object):
    """
    A object represents excel address mapping
    """
    def __init__(self, address: str = None, matrix: (tuple, list) = None):

        self._row, self._column, self._bottom, self._border = 0, 0, 0, 0

        if address:
            self.Parse(address)

        if matrix:
            if len(matrix) != 2:
                raise Exception()

            if self._row and self._row != matrix[0][0]:
                raise Exception()

            if self._column and self._column != matrix[0][1]:
                raise Exception()

        self._address = ''

    @property
    def Row(self):
        return self._row

    @property
    def Column(self):
        return self._column

    @property
    def Bottom(self):
        return self._bottom

    @property
    def Boundary(self):
        return self._border

    @property
    def Coord(self):
        return (self._row, self._column), (self._bottom, self._border)

    @property
    def Address(self):
        return self._address

    def Decode(self):
        pass

    def Encode(self):
        pass

    def Parse(self, address: str):
        pass

    def Dump(self) -> str:
        pass

    def IndexOf(self, coord: (tuple, list)) -> int:
        return 1

    def CoordOf(self, index: int) -> tuple:
        return (1, 1)

