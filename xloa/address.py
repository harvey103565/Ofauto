#! /usr/bin/python3
# -*- coding: utf-8 -*-


import re

from .xlerror import XlError


class Address(object):
    """
    A object represents excel address mapping
    """
    def __init__(self, address: str = None, matrix: (tuple, list) = None):

        self._row, self._column, self._bottom, self._border = 0, 0, 0, 0

        if all((address, matrix)):
            raise XlError()

        if address:
            self._address = address
            self._coords = Parse(address)

        if matrix:
            self._coords = matrix
            if all(isinstance(i, int) for i in matrix) and len(matrix) == 2:
                self._address = Encode(matrix)
            if all(isinstance(i, (list, tuple)) for i in matrix):
                self._address = Encode(*matrix)

        if not any((self._address, self._coords)):
            raise XlError()

    @property
    def Row(self):
        return self._coords[0][0]

    @property
    def Column(self):
        return self._coords[0][1]

    @property
    def Coord(self):
        return self._coords

    @property
    def Address(self):
        return self._address

    def Dump(self) -> str:
        return Encode(self._address)

    def IndexOf(self, coord: (tuple, list)) -> int:
        return 1

    def CoordOf(self, index: int) -> tuple:
        return (1, 1)


def Encode(*coords):
    if not all(len(coord) == 2 for coord in coords):
        raise XlError()

    address = ''
    for coord in coords:
        if address:
            address += ':'

        if coord[1]:
            address += Int2Letter26(coord[1])

        if coord[0]:
            address += str(coord[0])

    return address


def Parse(address: str) -> (tuple, tuple):
    return tuple(Decode(unit) for unit in address.split(':'))


def Decode(addr: str) -> tuple:
    units = re.findall(r'\d+|[A-Z]+', addr)

    if len(units) > 2:
        raise XlError()

    if units[0]:
        col = Letter2Int26(units[0])
    else:
        col = None

    if units[1]:
        row = int(units[1])
    else:
        row = None

    return row, col


__digits__ = (1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26)
__chrs__ = ('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q',
            'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z')

__dc__ = dict(zip(__chrs__, __digits__))


def Int2Letter26(n: int) -> str:
    if n < 1:
        raise XlError()

    address = ''
    while n > 0:
        n -= 1
        address = __chrs__[n % 26] + address
        n //= 26

    return address


def Letter2Int26(s: str) -> int:
    i = 0
    for c in s.upper():
        i = i * 26 + __dc__[c]

    return i

