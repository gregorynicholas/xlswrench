# coding: utf-8

import xlrd


class XlrdProxy(object):
    __slots__ = ("_book",)

    def __init__(self, book):
        self._book = book

    def get_number_of_sheets(self):
        return self._book.nsheets

    def get_sheet_name(self, sheet_index):
        return self._book.get_sheet_by_index(sheet_index).name

    def get_sheet_number_of_rows(self, sheet_index):
        return self._book.get_sheet_by_index(sheet_index).nrows

    def get_sheet_number_of_cols(self, sheet_index):
        return self._book.get_sheet_by_index(sheet_index).ncols

    def get_sheet_cell_value(self, sheet_index, row, col):
        return self._book.get_sheet_by_index(sheet_index).cell_value(row, col)


class Sheet(object):
    __slots__ = ("proxy",)

    def __init__(self, proxy):
        self.proxy = proxy

    @property
    def name(self):
        "Name of the sheet."
        return self.proxy.get_sheet_name()

    @property
    def rows(self):
        "Number of rows in the sheet."
        return self.proxy.get_sheet_number_of_rows()

    @property
    def cols(self):
        "Number of columns in the sheet."
        return self.proxy.get_sheet_number_of_cols()

    def cell(self, row, col):
        "Get the cell determined by `row` and `col` params."
        value = self.proxy.get_sheet_cell_value(row, col)
        return Cell(row, col, None, value)


class Row(object):
    __slots__ = ("index", "cells", "sheet")

    def __init__(self, index, cells, sheet):
        self.index = index
        self.cells = cells
        self.sheet = sheet

    def __getitem__(self, index):
        if isinstance(index, int):
            return self.cells[index]
        return [cell for cell in self.cells if cell.name == index]


class Cell(object):
    __slots__ = ("row", "col", "name", "value")

    def __init__(self, row, col, name, value):
        self.row = row
        self.col = col
        self.name = name
        self.value = value


def tour_xls(filename, callback, first_row_are_column_names=True, **callback_kwargs):
    """Tour all the records in an xls file calling a callback for each row.

    The callback receives a :class:`Row` instance as the first argument and additionally, the
    specified `callback_kwargs` as keyword arguments.

    :param filename: Filename of the xls file.
    :param callback: Callable that is going to be called for each row defined in the file.
    :param first_row_are_column_names: If `True` (default)
    :param callback_kwargs: When calling `callback`, also pass these parameters as keyword
    arguments.
    """
    with xlrd.open_workbook(xls) as book:
        book = XlrdProxy(book)
        for sheet in range(book.get_number_of_sheets()):

            if first_row_are_column_names:
                first_row = 1
                column_names = [book.get_sheet_cell_value(sheet, 0, col)
                                for col in range(book.get_sheet_number_of_cols(sheet))]
            else:
                first_row = 0
                column_names = [None] * book.get_sheet_number_of_cols(sheet)

            for row in range(first_row, book.get_sheet_number_of_rows(sheet)):
                cells = {}
                for col in range(sheet.cols):
                    value = sheet.value(row, col)
                    cells[column_names[col]] = Cell(row, col, column_names[col], value)

                callback(XlrdRowProxy(row_i, cells, Sheet(xlrd_sheet)),
                         callback, **callback_kwargs)
