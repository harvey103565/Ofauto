#! /usr/bin/python3
# -*- coding: utf-8 -*-


from win32com.client import Dispatch
from pythoncom import com_error
from functools import wraps
from timber import timber

from .charts import Chart
from .charts import Charts
from .macro import Macro
from .address import Address

from .interior import ColorIndex
from .interior import int_to_rgb
from .interior import rgb_to_int

from .xlerror import XlError


def try_each(*tars):

    def decoration(func):

        @wraps(func)
        def wrapper(self, *args, **kwargv):
            for t in tars:
                try:
                    com_object = func(self, t)
                    timber.info('com object {0} created successfully'.format(t))
                    return com_object
                except com_error as err:
                    timber.error(err)

        return wrapper

    return decoration


class App(object):
    """
    The class that works as the agent between client and windows com object: Application.
    The instance of this
    """

    def __init__(self, impl=None):
        if impl:
            self._impl = impl
        else:
            self._impl = self.Connect()

    @try_each('Excel.Application', 'Ket.Application')
    def Connect(self, cls_id):
        return Dispatch(cls_id)

    @property
    def Api(self):
        """
        Return com object
        :return: com object
        """
        return self._impl

    @property
    def ActiveCell(self):
        return self.Api.ActiveCell

    @property
    def ActiveWorkbook(self):
        return Book(self.Api.ActiveWorkbook)

    @property
    def ActiveSheet(self):
        return Sheet(self.Api.ActiveSheet)

    @property
    def Workbooks(self):
        return Books(self.Api.Workbooks)

    @property
    def Worksheets(self):
        return Sheets(self.Api.Worksheets)

    @property
    def Visible(self):
        return self.Api.Visible

    @Visible.setter
    def Visible(self, visibility: bool = True):
        """
        Set the visibility of Excel application
        :param visibility: bool, whether this application is visible
        :return: N/A
        """
        self.Api.Visible = visibility

    def Macro(self, name):
        """
        Runs a Sub or Function in Excel VBA that are not part of a specific workbook but e.g. are part of an add-in.
        :param name: Name of Sub or Function with or without module name, e.g. ``'Module1.MyMacro'`` or ``'MyMacro'``
        :return: Macro results
        """
        return Macro(self, name)


class Book(object):
    """
    CLass that represents a Workbook object
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

    @property
    def Application(self) -> App:
        """
        The Application of Workbook
        :return: App instance
        """
        return App(self.Api.Application)

    @property
    def Name(self) -> str:
        """
        The file name of Workbook
        :return: Workbook name
        """""
        return self.Api.Name

    @property
    def FullName(self) -> str:
        """
        The full path name of Workbook
        :return: Workbook path/name
        """
        return self.Api.FullName

    @property
    def Worksheets(self):
        """
        Sheet collection of Workbook
        :return: Sheets instance
        """
        return Sheets(self.Api.Worksheets)

    @property
    def ActiveSheet(self):
        """
        Active sheet of Workbook
        :return: Active sheet instance
        """
        return Sheet(self.Api.ActiveSheet)

    @property
    def Charts(self) -> Charts:
        """
        Charts collection of Workbook
        :return: Charts instance
        """
        return Charts(self.Api.Charts)

    @property
    def ActiveChart(self) -> Chart:
        """
        Active chart of Workbook
        :return: Chart instance
        """
        return Chart(self.Api.ActiveChart)


class Books(object):
    """
    Class that represents Collection of all Workbooks
    """
    def __init__(self, impl):
        self._impl = impl

    def __getitem__(self, item) -> Book:
        """
        Get an Workbook instance directly by index or id(name str)
        :param item: index(int) or id(name str) of the target Workbook
        :return: Workbook instance
        """
        if isinstance(item, int) or isinstance(item, str):
            return Book(self.Api.Item(item))
        else:
            raise XlError('Use <int> or <str> as subscript.')

    def __iter__(self) -> iter:
        """
        Iterate through Workbooks
        :return: Generator for iteration
        """
        for i in range(self.Count):
            yield Book(self[i])

    def __len__(self):
        """
        Return the number of Workbooks in this collection
        :return: count of workbooks
        """
        return self.Count

    @property
    def Api(self):
        """
        Return com object
        :return: com object
        """
        return self._impl

    @property
    def Count(self) -> int:
        return self.Api.Count

    @property
    def Application(self) -> App:
        return App(self.Api.Application)

    def Add(self) -> Book:
        """
        Add a workbook
        :return:
        """
        book = Book(self.Api.Add())

        if not self.Application.Visible:
            self.Application.Visible = True

        return book

    def Close(self):
        """
        Close all opened Workbooks The Excel will popup a save dialogue when there are Workbooks modified
        """
        self.Api.Close()

    def Open(self, file_name: str) -> Book:
        return Book(self.Open(file_name))


class Range(object):
    """
    Range object
    """
    def __init__(self, impl, item_step: tuple = None):
        self._impl = impl
        self._row_count = impl.Rows.Count
        self._column_count = impl.Columns.Count
        self._item_step = item_step

    @property
    def Api(self):
        """
        Return com object wrapped by Range object
        :return: Com instance
        """
        return self._impl

    @property
    def Worksheet(self):
        return Sheet(self.Api.Worksheet)

    @property
    def Address(self) -> str:
        """
        Only used cells in Worksheet, this is included in a rectangle area
        :return:
        """
        return self._impl.Address

    @property
    def Row(self):
        return self.Api.Row

    @property
    def Column(self):
        return self.Api.Column

    @property
    def RowCount(self):
        return self._row_count

    @property
    def ColumnCount(self):
        return self._column_count

    @property
    def Rows(self):
        """
        All Rows in Worksheet
        :return:
        """
        return Rows(self._impl.Rows)

    @property
    def Columns(self):
        """
        All Columns in Worksheet
        :return:
        """
        return Columns(self._impl.Columns)

    @property
    def Count(self):
        return self._impl.Count

    @property
    def Text(self):
        """
        Return Text of a cell in the Range
        :return: Value
        """
        return self.Api.Text

    @property
    def Value(self):
        """
        Return value of cells in the Range
        :return: Value
        """
        return self.Api.Value

    @Value.setter
    def Value(self, value):
        """
        Set Value of Range
        :return: NA
        """
        self.Api.Value = value

    def __call__(self, *args, **kwargs):
        """
        Returns a Range object that represents the cells in the specified Range. By apply __call__() method to
        Cells property, which returns this Range object itself, a new Range is generated.
            Cells(1)
            Cells(1, 2)
            Cells((1, 2)) or Cells([1, 2])
            Cells((1, 2), [3, 4]) or Cells([1, 2], (3, 4))
            rng[(1, 1), (2, 2)], rng[(1, 1): (2, 2)]    - cells from [1st row, 1st col] to [2nd row, 2nd col]
        :param args:
            int - the index of cells in order from left to right and then next row
            tuple or list - the index tuple with length of 2, and in form of (row, column)/[row, column]
        :param kwargs: Not used yet
        :return: Range Object
        """
        """
        Return Range object specified by matrix: (1, 1), ((1, 1)) or ((1, 1), (2, 2))
        :param matrix: matrix contains the coordinates in tuple or 2d tuple
        :return: Range object
        """
        if len(args) == 0:
            return Range(self.Api.Cells)

        if len(args) > 2:
            raise XlError('Use (int, int)[, (int, int)] matrix syntax.')

        if len(args) == 1 or all(isinstance(k, int) for k in args):
            return Range(self.Api.Cells(*args))

        if any(not isinstance(k, (list, tuple)) for k in args):
            raise XlError()

        return Range(self.Api.Range(Address(matrix=args).Address))

    def __getitem__(self, indices):
        """
        Returns a Range object that represents a Range at an offset to the specified Range. Either by index or offset
        address or index tuple.
            rng[1]                                  - first cell
            rng[1, 2], rng[(1, 2)]                  - cell at [1st row, 2nd col]
            rng[1: 2]                               - cells include 1st to 2nd
            rng[(1, 1): (2, 2)]                     - cells from (1, 1) to (2, 2)
            rng["A1"], rng["A:B"], rng["A1:B2"]     - Cell(s) by excel address
        :param indices:
            int - the index of cells in order from left to right and then next row
            str - the address in string, Column index given in letters
            tuple or list - the index tuple with length of 2, and in form of (row, column)/[row, column]
        :return: Range Object
        """
        if isinstance(indices, int):
            return self.Item(indices)

        if isinstance(indices, str):
            return self.Range(indices)

        if isinstance(indices, slice):
            if indices.step:
                raise XlError('Slice step can not be applied to range object.')

            if all(isinstance(p, (tuple, list)) for p in (indices.start, indices.stop) if p):
                if indices.start:
                    ep1 = Range.Coordinates(indices.start, row_count=self.RowCount, column_count=self.ColumnCount)
                else:
                    ep1 = (1, 1)

                if indices.stop:
                    ep2 = Range.Coordinates(indices.stop, row_count=self.RowCount, column_count=self.ColumnCount)
                else:
                    ep2 = (self.RowCount, self.ColumnCount)

                if ep1[0] + ep2[0] - 1 > self.RowCount:
                    raise XlError('Index of row out of range.')

                if ep1[1] + ep2[1] - 1 > self.ColumnCount:
                    raise XlError('Index of column out of range.')

                matrix = ((min(ep1[0], ep2[0]), min(ep1[1], ep2[1])), (max(ep1[0], ep2[0]), max(ep1[1], ep2[1])))

                return Range(self.Api.Range(Address(matrix=matrix).Address))

        raise XlError('Use int, str or slice(tuple(int, int): tuple(int, int)) as subscriber.')

    def __len__(self):
        """
        Return the Count
        :return:
        """
        return self.Count

    def __iter__(self):
        """
        Return an iterator for cells in Range
        :return: Range Object
        """
        for i in range(self.Count):
            yield Range(self.Api.Item(i + 1))

    @staticmethod
    def Coordinates(*coord, row: int = 1, column: int = 1, row_count: int = 0, column_count: int = 0) -> tuple:
        if not any(coord):
            return row, column

        if len(coord) == 1:
            if isinstance(coord[0], (tuple, list)) and len(coord[0]) == 2:
                row, column = coord[0]
            else:
                raise XlError('Coordinate should be 2 dimension vector.')

        if len(coord) == 2:
            if all(isinstance(c, int) for c in coord):
                row, column = coord
            else:
                raise XlError('Coordinate should be 2 dimension vector.')

        if len(coord) > 2:
            raise XlError('Coordinate should be 2 int numbers standalone or in a tuple.')

        if row_count:
            if row < -row_count or row > row_count:
                raise XlError('row coordinate should be in +/- [1, row_count]')

            if -row_count <= row < 0:
                row += (row_count + 1)

        if column_count:
            if column < -column_count or column > column_count:
                raise XlError('row coordinate should be in +/- [1, row_count]')

            if -column_count <= column < 0:
                column += (column_count + 1)

        return row, column

    def Range(self, *endpoints):
        """
        Returns a Range object that represents a Range at an offset to the specified Range. Either by index or offset
            address or index tuple. Example:
            Range(rng, rng)
            Rnage('A1:B2')
        :param endpoints: accept a list of endpoints, either in address format or a couple of Range objects
            str - the address in string, Column index given in letters
            tuple or list - the index tuple with length of 2, with 2 Range objects indicates the endpoints
        :return: Range Object
        """
        if all(isinstance(ep, Range) for ep in endpoints):
            if len(endpoints) == 2:
                return Range(self.Api.Range(endpoints[0].Api, endpoints[1].Api))
            else:
                raise XlError('Must be 2 Range objects as endpoints.')

        return Range(self.Api.Range(*endpoints))

    def Cells(self, *coords):
        """
        Returns a Range object that represents a cell in the specified Range. This cell could be accessed by following
        ways:
            Cells(1)    - The 1st cell in Range
            Cells(1, 2) - The cell at row 1 Colmumn 2
        :param coords: the coordinates of the cell to access, in order from left to right, and then down
            Range.Cells(1) returns the upper-left cell in the Range.
            Range.Cells(2) returns the cell immediately to the right of the upper-left cell.
        :return: Range Object it self
        """
        if len(coords) == 0:
            return Range(self.Api.Cells)

        if len(coords) > 2:
            raise XlError('Use (int, int)[, (int, int)] matrix syntax.')

        if any(not isinstance(k, int) for k in coords):
            raise XlError('Cells method only accepts int as input.')

        return Range(self.Api.Cells(*coords))

    def Item(self, *coords):
        """
        Returns a Range object that represents a Range at an offset to the specified Range, Note the basic unit to
        count is not limited to single cell, it could also be a row, a column, depending on how this Range object is
        generated. *** This is different behavior compare with __call__ method.
            Item(1)
            Item(1, 2)
        :param coords: the coordinates of the cell to access, in order from left to right, and then down
            Range.Cells(1) returns the upper-left cell in the Range.
            Range.Cells(2) returns the cell immediately to the right of the upper-left cell.
        :return: Range Object
        """
        if len(coords) == 0:
            raise XlError()

        if len(coords) > 2:
            raise XlError('Use (int, int)[, (int, int)] matrix syntax.')

        if not all(isinstance(k, int) for k in coords):
            raise XlError('')

        return Range(self.Api.Item(*coords))

    @property
    def Color(self):
        if self.Api.Interior.ColorIndex == ColorIndex.xlColorIndexNone:
            return None
        else:
            return int_to_rgb(self.Api.Interior.Color)

    @Color.setter
    def Color(self, color_or_rgb):
        if color_or_rgb is None:
            self.Api.Interior.ColorIndex = ColorIndex.xlColorIndexNone
        elif isinstance(color_or_rgb, int):
            self.Api.Interior.Color = color_or_rgb
        else:
            self.Api.Interior.Color = rgb_to_int(color_or_rgb)


class Sheet(Range):
    """
    Class represents Worksheet object in Workbook
    """
    @property
    def Application(self) -> App:
        """
        The Application of Workbook
        :return: App instance
        """
        return App(self.Api.Application)

    @property
    def Name(self) -> str:
        """
        The file name of Workbook
        :return: Workbook name
        """""
        return self.Api.Name

    @Name.setter
    def Name(self, name: str):
        """
        Set name for worksheet
        :param name: str, the new name
        :return: None
        """
        self.Api.Name = name

    @property
    def UsedRange(self):
        """
        Only used cells in Worksheet, this is included in a rectangle area
        :return:
        """
        return Range(self.Api.UsedRange)


class Sheets(object):
    """
    Class represents Collection of all Worksheets in Workbook
    """
    def __init__(self, impl):
        self._impl = impl

    def __getitem__(self, item):
        """
        Get an Worksheet instance directly by index or id(name str)
        :param item: index(int) or id(name str) of the target Worksheet
        :return: Worksheet instance
        """
        if isinstance(item, int) or isinstance(item, str):
            return Sheet(self.Api[item])
        else:
            raise XlError('Use <int> or <str> as subscript.')

    def __iter__(self):
        """
        Iterate through Workbooks
        :return: Generator for iteration
        """
        for i in range(self.Count):
            yield Sheet(self[i])

    def __len__(self):
        """
        Return number of Worksheets in this collection
        :return: count of worksheets
        """
        return self.Count

    @property
    def Api(self):
        """
        Return com object
        :return: com object
        """
        return self._impl

    @property
    def Count(self):
        return self.Api.Count

    @property
    def Application(self):
        return App(self.Api.Application)

    def Add(self, name: str = None):
        """
        Add a worksheet to collection with specified name
        :param name:
        :return:
        """
        sheet = Sheet(self.Api.Add())
        if name:
            sheet.Name = name
        return sheet

    def Close(self):
        self.Api.Close()

    def Open(self, file_name: str) -> Book:
        return Book(self.Open(file_name))


class Rows(Range):
    """
    Object represents a collection of Rows in Range
    """
    def __init__(self, impl):
        super().__init__(impl)

    def __call__(self, *args, **kwargs):
        if len(args) == 1 and isinstance(args[0], int):
            return Rows(self.Api.Rows(args[0]))

        return super().__call__(*args)

    def __iter__(self):
        """
        Return an iterator for cells in Range
        :return: Range Object
        """
        for i in range(self.Count):
            yield Rows(self.Api.Item(i + 1))

    def __getitem__(self, indices):
        """
        Return a Range object represents the specified range in worksheet. Depending on the different input parameters,
        Rows or Range object may returned. For example:
            rows[1]                 - will return the 1st row in range
            rows[2: 5]              - will return a Rows object including a subset of rows, from row 2 to row 5
            rows['2:5']             - will also return a subset of rows
            rows[(1, 1), (2, 2)[    - will return a Range object represents subset cells of Range(Rows)
                                      in the rectangle: row 1 col 1 to row 2 col 2
            rows['A1:B2']           - will return a Range object represents subset cells of Range(Rows)
                                      in the rectangle: row 1 col 1 to row 2 col 2
        :param indices: the indices used to specify the range or rows
        :return:
        """
        if isinstance(indices, int):
            return self.Item(indices)

        if isinstance(indices, str):
            addr = Address(addr=indices)
            if addr.IsRow:
                if addr.Row + addr.RowCount > self.RowCount:
                    raise XlError('Index of row out of range.')

                return Rows(self.Api.Rows(str))

            return self.Range(indices)

        if isinstance(indices, slice):
            if any(not isinstance(p, int) for p in (indices.start, indices.stop) if p):
                return super().__getitem__(indices)

            if indices.step:
                raise XlError('Slice step can not be applied to range object.')

            if indices.stop > self.RowCount:
                raise XlError('Index of row out of range.')

            matrix = ((indices.start or 1, 0), (indices.stop or self.Count, 0))
            return Rows(self.Api.Range(Address(matrix=matrix).Address).Rows)

        raise XlError('')

    @property
    def Value(self):
        """
        Return value of cells in the Range
        :return: Value
        """
        return self.Api.Rows.Value

    @Value.setter
    def Value(self, value):
        """
        Set Value of Range
        :return: NA
        """
        self.Api.Rows.Value = value

    def Item(self, item):
        return Rows(self.Api.Rows(item))


class Columns(Range):
    """
    Object represents a collection of Rows in Range
    """

    def __init__(self, impl):
        super().__init__(impl)

    def __call__(self, *args, **kwargs):
        if len(args) == 1 and isinstance(args[0], int):
            return Columns(self.Api.Columns(args[0]))

        return super().__call__(*args)

    def __iter__(self):
        """
        Return an iterator for cells in Range
        :return: Range Object
        """
        for i in range(self.Count):
            yield Columns(self.Api.Item(i + 1))

    def __getitem__(self, indices):
        """
        Return a Range object represents the specified range in worksheet. Depending on the different input parameters,
        Columns or Range object may returned. For example:
            columns[1]                  - will return the 1st columns in range
            columns[2: 5]               - will return a Rows object including a subset of columns,
                                          from column 2 to column 5
            columns['2:5']              - will also return a subset of rows
            columns[(1, 1), (2, 2)[     - will return a Range object represents subset cells of Range(Columns)
                                          in the rectangle: row 1 col 1 to row 2 col 2
            columns['A1:B2']            - will return a Range object represents subset cells of Range(Columns)
                                          in the rectangle: row 1 col 1 to row 2 col 2
        :param indices: the indices used to specify the range or columns
        :return:
        """
        if isinstance(indices, int):
            return self.Item(indices)

        if isinstance(indices, str):
            addr = Address(addr=indices)
            if addr.IsColumn:
                if addr.Column + addr.ColumnCount > self.ColumnCount:
                    raise XlError('Index of column out of range.')

                return Columns(self.Api.Columns(str))

            return self.Range(indices)

        if isinstance(indices, slice):
            if any(not isinstance(p, int) for p in (indices.start, indices.stop) if p):
                return super().__getitem__(indices)

            if indices.step:
                raise XlError('Slice step can not be applied to range object.')

            if indices.stop > self.RowCount:
                raise XlError('Index of column out of range.')

            matrix = ((0, indices.start or 1), (0, indices.stop or self.Count))
            return Columns(self.Api.Range(Address(matrix=matrix).Address).Columns)

        raise XlError('')

    @property
    def Value(self):
        """
        Return value of cells in the Range
        :return: Value
        """
        return self.Api.Columns.Value

    @Value.setter
    def Value(self, value):
        """
        Set Value of Range
        :return: NA
        """
        self.Api.Columns.Value = value

    def Item(self, item):
        return Columns(self.Api.Columns(item))
