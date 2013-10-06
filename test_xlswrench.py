# coding: utf-8

import mock
import xlswrench


class TestCell:
    def setup(self):
        self.cell = xlswrench.Cell(1, 2, None, "cell value")

    def test_row(self):
        assert self.cell.row == 1

    def test_col(self):
        assert self.cell.col == 2

    def test_name(self):
        assert self.cell.name is None

    def test_value(self):
        assert "cell value" == self.cell.value


class TestRow:
    def setup(self):
        self.cells = [xlswrench.Cell(1, 2, "cell name", "cell value")]
        self.sheet = mock.Mock()
        self.row = xlswrench.Row(3, self.cells, self.sheet)

    def test_index(self):
        assert self.row.index == 3

    def test_cells(self):
        assert self.row.cells == self.cells

    def test_sheet(self):
        assert self.row.sheet == self.sheet

    def test_access_a_cell_by_index(self):
        assert self.row[0] == self.cells[0]

    def test_access_all_cells_with_a_name(self):
        assert self.row["cell name"] == [self.cells[0]]

    def test_empty_list_if_no_cells_match_the_name(self):
        assert self.row["something"] == []


class TestSheet:
    def setup(self):
        self.sheet = xlswrench.Sheet(proxy=mock.Mock())

    def test_name(self):
        self.sheet.name
        self.sheet.proxy.get_sheet_name.assert_called_once_with()

    def test_rows(self):
        self.sheet.rows
        self.sheet.proxy.get_sheet_number_of_rows.assert_called_once_with()

    def test_cols(self):
        self.sheet.cols
        self.sheet.proxy.get_sheet_number_of_cols.assert_called_once_with()

    def test_cell(self):
        self.sheet.cell(0, 1)
        self.sheet.proxy.get_sheet_cell_value.assert_called_once_with(0, 1)

    def test_cell_instance(self):
        assert isinstance(self.sheet.cell(0, 1), xlswrench.Cell)


class TestXlrdProxy:
    def setup(self):
        sheet_mock = mock.Mock()
        sheet_mock.name = "sheet name"
        sheet_mock.nrows = 3
        sheet_mock.ncols = 4
        sheet_mock.cell_value = mock.Mock(return_value="cell value")

        book_mock = mock.Mock()
        book_mock.nsheets = 1
        book_mock.get_sheet_by_index = mock.Mock(return_value=sheet_mock)

        self.proxy = xlswrench.XlrdProxy(book_mock)

    def test_get_number_of_sheets(self):
        assert self.proxy.get_number_of_sheets() == 1

    def test_get_name(self):
        name = self.proxy.get_sheet_name(0)
        self.proxy._book.get_sheet_by_index.assert_called_once_with(0)
        assert name == "sheet name"

    def test_get_sheet_number_of_rows(self):
        number_of_rows = self.proxy.get_sheet_number_of_rows(0)
        self.proxy._book.get_sheet_by_index.assert_called_once_with(0)
        assert number_of_rows == 3

    def test_get_sheet_number_of_cols(self):
        number_of_columns = self.proxy.get_sheet_number_of_cols(0)
        self.proxy._book.get_sheet_by_index.assert_called_once_with(0)
        assert number_of_columns == 4

    def test_get_sheet_cell_value(self):
        value = self.proxy.get_sheet_cell_value(0, 0, 2)
        self.proxy._book.get_sheet_by_index.assert_called_once_with(0)
        self.proxy._book.get_sheet_by_index.return_value.cell_value.assert_called_once_with(0, 2)
        assert value == "cell value"


class TestTourXls:
    def setup(self):
        pass
