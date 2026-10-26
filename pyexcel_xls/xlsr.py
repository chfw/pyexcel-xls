"""
    pyexcel_xlsr
    ~~~~~~~~~~~~~~~~~~~

    The lower level xls/xlsm file format handler using xlrd

    :copyright: (c) 2016-2021 by Onni Software Ltd
    :license: New BSD License
"""

import datetime

import xlrd
from pyexcel_io.service import has_no_digits_in_float
from pyexcel_io.plugin_api import ISheet, IReader

XLS_KEYWORDS = [
    "filename",
    "logfile",
    "verbosity",
    "use_mmap",
    "file_contents",
    "encoding_override",
    "formatting_info",
    "on_demand",
    "ragged_rows",
]
DEFAULT_ERROR_VALUE = "#N/A"


class MergedCell(object):
    def __init__(self, row_low, row_high, column_low, column_high):
        self.__rl = row_low
        self.__rh = row_high
        self.__cl = column_low
        self.__ch = column_high
        self.value = None

    def register_cells(self, registry):
        for rowx in range(self.__rl, self.__rh):
            for colx in range(self.__cl, self.__ch):
                key = "%s-%s" % (rowx, colx)
                registry[key] = self


class XLSheet(ISheet):
    """
    xls, xlsx, xlsm sheet reader

    Currently only support first sheet in the file
    """

    def __init__(self, sheet, auto_detect_int=True, date_mode=0, **keywords):
        self.__auto_detect_int = auto_detect_int
        self.__hidden_cols = []
        self.__hidden_rows = []
        self.__merged_cells = {}
        self._book_date_mode = date_mode
        self.xls_sheet = sheet
        self._keywords = keywords
        if keywords.get("detect_merged_cells") is True:
            for merged_cell_ranges in sheet.merged_cells:
                merged_cells = MergedCell(*merged_cell_ranges)
                merged_cells.register_cells(self.__merged_cells)
        if keywords.get("skip_hidden_row_and_column") is True:
            for col_index, info in self.xls_sheet.colinfo_map.items():
                if info.hidden == 1:
                    self.__hidden_cols.append(col_index)
            for row_index, info in self.xls_sheet.rowinfo_map.items():
                if info.hidden == 1:
                    self.__hidden_rows.append(row_index)

    @property
    def name(self):
        return self.xls_sheet.name

    def row_iterator(self):
        number_of_rows = self.xls_sheet.nrows - len(self.__hidden_rows)
        return range(number_of_rows)

    def column_iterator(self, row):
        number_of_columns = self.xls_sheet.ncols - len(self.__hidden_cols)
        for column in range(number_of_columns):
            yield self.cell_value(row, column)

    def cell_value(self, row, column):
        """
        Random access to the xls cells
        """
        if self._keywords.get("skip_hidden_row_and_column") is True:
            row, column = self._offset_hidden_indices(row, column)
        cell_type = self.xls_sheet.cell_type(row, column)
        value = self.xls_sheet.cell_value(row, column)

        if cell_type == xlrd.XL_CELL_DATE:
            value = xldate_to_python_date(value, self._book_date_mode)
        elif cell_type == xlrd.XL_CELL_NUMBER and self.__auto_detect_int:
            if has_no_digits_in_float(value):
                value = int(value)
        elif cell_type == xlrd.XL_CELL_ERROR:
            value = DEFAULT_ERROR_VALUE

        if self.__merged_cells:
            merged_cell = self.__merged_cells.get("%s-%s" % (row, column))
            if merged_cell:
                if merged_cell.value:
                    value = merged_cell.value
                else:
                    merged_cell.value = value
        return value

    def _offset_hidden_indices(self, row, column):
        row = calculate_offsets(row, self.__hidden_rows)
        column = calculate_offsets(column, self.__hidden_cols)
        return row, column


def calculate_offsets(incoming_index, hidden_indices):
    offset = 0
    for index in hidden_indices:
        if index <= (incoming_index + offset):
            offset += 1
    return incoming_index + offset


class XLSReader(IReader):
    """
    XLSBook reader

    It reads xls, xlsm, xlsx work book
    """

    def __init__(self, file_type, **keywords):
        self.__skip_hidden_sheets = keywords.get("skip_hidden_sheets", True)
        self.__skip_hidden_row_column = keywords.get(
            "skip_hidden_row_and_column", True
        )
        self.__detect_merged_cells = keywords.get("detect_merged_cells", False)
        self._keywords = keywords
        self._date_mode = self._keywords.pop("date_mode", None)
        xlrd_params = self._extract_xlrd_params()
        if self.__skip_hidden_row_column and file_type == "xls":
            xlrd_params["formatting_info"] = True
        if self.__detect_merged_cells:
            xlrd_params["formatting_info"] = True

        self.content_array = []
        self.xls_book = self.get_xls_book(**xlrd_params)
        for sheet in self.xls_book.sheets():
            if self.__skip_hidden_sheets and sheet.visibility != 0:
                continue
            self.content_array.append(sheet)

    def read_sheet(self, index):
        native_sheet = self.content_array[index]
        date_mode = (
            self.xls_book.datemode
            if self._date_mode is None
            else self._date_mode
        )
        sheet = XLSheet(native_sheet, date_mode=date_mode, **self._keywords)
        return sheet

    def close(self):
        if self.xls_book:
            self.xls_book.release_resources()
            self.xls_book = None

    def get_xls_book(self, **xlrd_params):
        xls_book = xlrd.open_workbook(**xlrd_params)
        return xls_book

    def _extract_xlrd_params(self):
        params = {}
        if self._keywords is not None:
            for key in list(self._keywords.keys()):
                if key in XLS_KEYWORDS:
                    params[key] = self._keywords.pop(key)
        return params


class XLSInFile(XLSReader):
    def __init__(self, file_name, file_type, **keywords):
        super().__init__(file_type, filename=file_name, **keywords)


class XLSInContent(XLSReader):
    def __init__(self, file_content, file_type, **keywords):
        super().__init__(file_type, file_contents=file_content, **keywords)


class XLSInMemory(XLSReader):
    def __init__(self, file_stream, file_type, **keywords):
        file_stream.seek(0)
        super().__init__(
            file_type, file_contents=file_stream.read(), **keywords
        )


def xldate_to_python_date(value, date_mode):
    """
    convert xl date to python date
    """
    date_tuple = xlrd.xldate_as_tuple(value, date_mode)

    ret = None
    if date_tuple == (0, 0, 0, 0, 0, 0):
        ret = datetime.datetime(1900, 1, 1, 0, 0, 0)
    elif date_tuple[0:3] == (0, 0, 0):
        ret = datetime.time(date_tuple[3], date_tuple[4], date_tuple[5])
    elif date_tuple[3:6] == (0, 0, 0):
        ret = datetime.date(date_tuple[0], date_tuple[1], date_tuple[2])
    else:
        ret = datetime.datetime(
            date_tuple[0],
            date_tuple[1],
            date_tuple[2],
            date_tuple[3],
            date_tuple[4],
            date_tuple[5],
        )
    return ret
