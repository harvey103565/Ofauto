#! /usr/bin/python3
# -*- coding: utf-8 -*-


import re

from .xlerror import XlError


class Address(object):
    """
    An class that represents relative address expression
    """
    def __init__(self, addr: str = None, matrix: tuple or list = None):
        if all((addr, matrix)):
            raise XlError('Do not set both add-ref and matrix.')

        if addr:
            self._addr = addr.upper()
            self._matrix = Decode(self._addr)

        if matrix:
            self._matrix = matrix

            if all(isinstance(i, int) for i in matrix) and len(matrix) == 2:
                self._addr = Encode(matrix)

            if all(isinstance(i, (list, tuple)) for i in matrix):
                self._addr = Encode(*matrix)

        if not all((self._addr, self._matrix, self._matrix[0])):
            raise XlError('No valid add-ref or matrix assigned.')

    @property
    def Row(self):
        return self._matrix[0][0] or 1

    @property
    def Column(self):
        return self._matrix[0][1] or 1

    @property
    def RowCount(self):
        if len(self._matrix) == 1:
            return 1

        return (self._matrix[1][0] or 1048576) - self.Row + 1

    @property
    def ColumnCount(self):
        if len(self._matrix) == 1:
            return 1

        return (self._matrix[1][1] or 16384) - self.Column + 1

    @property
    def Matrix(self):
        return self._matrix

    @property
    def Address(self):
        return self._addr

    @property
    def IsRow(self):
        return not self._matrix[0][1] or not (self._matrix[1] and self._matrix[1][1])

    @property
    def IsColumn(self):
        return not self._matrix[0][0] or not (self._matrix[1] and self._matrix[1][0])

    def IndexOf(self, coord: tuple, abs: bool = False) -> int:
        if len(coord) == 2 and all(isinstance(c, int) for c in coord):
            row, col = coord[0], coord[1]
            if abs:
                row, col = row - self.Row + 1, col - self.Column + 1

            return (row - 1) * self.ColumnCount + col

    def CoordOf(self, index: int) -> tuple:
        if index < 1:
            raise XlError('Index can not be less than 1.')

        return (index - 1) // self.ColumnCount + 1, (index - 1) % self.ColumnCount + 1


def Encode(*coords):
    if any(len(coord) != 2 for coord in coords):
        raise XlError('Must offer both row and column index for coordinates.')

    address = ''
    for coord in coords:
        if address:
            address += ':'

        if coord[1]:
            address += Int2Letter26(coord[1])

        if coord[0]:
            address += str(coord[0])

    return address


def Decode(address: str) -> (tuple, tuple):
    return tuple(Parse(unit) for unit in address.split(':'))


def Parse(addr: str) -> tuple:
    units = re.findall(r'\d+|[A-Z]+', addr.upper())

    if len(units) > 2:
        raise XlError()

    row, col = None, None
    for unit in units:
        if unit.isdigit():
            row = int(unit)

        if unit.isalpha():
            col = Letter2Int26(unit)

    return row, col


def Int2Letter26(n: int) -> str:
    if n < 1:
        raise XlError('Coordinate in address matrix could not be lower than 1.')

    address = ''
    while n > 0:
        n -= 1
        address = __chars__[n % 26] + address
        n //= 26

    return address


def Letter2Int26(s: str) -> int:
    i = 0
    for c in s.upper():
        i = i * 26 + __dc__[c]

    return i


__digits__ = (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26)
__chars__ = ('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q',
            'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z')

__dc__ = dict(zip(__chars__, __digits__))
