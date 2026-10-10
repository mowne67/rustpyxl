"""Streaming XLSX operations. This is an explicit subset of openpyxl's API."""
from itertools import islice
from os import fspath
from weakref import WeakSet

from ._native import NativeWorkbook, NativeWriter

__version__ = "0.1.0"
__all__ = ["Workbook", "load_workbook"]


def _batch_size(value):
    if not isinstance(value, int) or isinstance(value, bool) or not 1 <= value <= 65536:
        raise ValueError("batch_size must be an integer in 1..65536")
    return value


class ReadOnlyWorksheet:
    def __init__(self, workbook, index, title):
        self._workbook, self._index, self.title = workbook, index, title

    def iter_rows(self, min_row=1, max_row=None, min_col=1, max_col=None,
                  values_only=False, *, batch_size=1024):
        if not values_only:
            raise NotImplementedError("this milestone supports values_only=True")
        _batch_size(batch_size)
        for name, value, limit in (("min_row", min_row, 1048576), ("max_row", max_row, 1048576),
                                   ("min_col", min_col, 16384), ("max_col", max_col, 16384)):
            if value is not None and (not isinstance(value, int) or isinstance(value, bool) or not 1 <= value <= limit):
                raise ValueError(f"{name} must be in 1..{limit}")
        if (max_row is not None and max_row < min_row) or (max_col is not None and max_col < min_col):
            raise ValueError("maximum bounds must be at least minimum bounds")
        self._workbook._ensure_open()
        native = self._workbook._native.rows(self._index, self._workbook._data_only)
        self._workbook._iterators.add(native)

        def rows():
            number = 0
            width = max_col
            while max_row is None or number < max_row:
                self._workbook._ensure_open()
                count = batch_size if max_row is None else min(batch_size, max_row - number)
                batch = native.batch(count)
                if not batch:
                    break
                for row in batch:
                    self._workbook._ensure_open()
                    number += 1
                    if width is None:
                        width = len(row)
                    if number < min_row:
                        continue
                    end = max_col if max_col is not None else len(row)
                    if len(row) < end:
                        row += (None,) * (end - len(row))
                    yield row[min_col - 1:end]
        return rows()

    @property
    def values(self):
        return self.iter_rows(values_only=True)


class ReadOnlyWorkbook:
    def __init__(self, filename, data_only):
        self._native = NativeWorkbook(fspath(filename))
        self._data_only = data_only
        self._closed = False
        self._iterators = WeakSet()
        self.worksheets = [ReadOnlyWorksheet(self, i, title) for i, title in enumerate(self._native.sheetnames)]

    def _ensure_open(self):
        if self._closed:
            raise ValueError("workbook is closed")

    @property
    def sheetnames(self):
        return [s.title for s in self.worksheets]

    @property
    def active(self):
        return self.worksheets[0] if self.worksheets else None

    def __getitem__(self, title):
        self._ensure_open()
        for sheet in self.worksheets:
            if sheet.title == title:
                return sheet
        raise KeyError(title)

    def close(self):
        self._closed = True
        for iterator in self._iterators:
            iterator.close()
        self._iterators.clear()
        self._native = None

    def __enter__(self):
        self._ensure_open()
        return self

    def __exit__(self, *args):
        self.close()


def load_workbook(filename, read_only=False, data_only=False, **kwargs):
    if read_only is not True:
        raise NotImplementedError("this milestone requires read_only=True")
    if kwargs:
        raise NotImplementedError(f"unsupported options: {', '.join(kwargs)}")
    return ReadOnlyWorkbook(filename, data_only)


class WriteOnlyWorksheet:
    def __init__(self, workbook, index, title):
        self._workbook, self._index, self.title = workbook, index, title

    def append(self, row):
        self._workbook._ensure_open()
        self._workbook._native.append_rows(self._index, [row])

    def append_rows(self, rows, *, batch_size=1024):
        """Consume an iterable in bounded batches to amortize Python/Rust calls.

        If an error occurs, preceding rows/batches may already be appended.
        """
        _batch_size(batch_size)
        self._workbook._ensure_open()
        iterator = iter(rows)
        while batch := list(islice(iterator, batch_size)):
            self._workbook._ensure_open()
            self._workbook._native.append_rows(self._index, batch)


class Workbook:
    def __init__(self, write_only=False):
        if write_only is not True:
            raise NotImplementedError("this milestone requires write_only=True")
        self._native = NativeWriter()
        self._closed = False
        self.worksheets = []

    def _ensure_open(self):
        if self._closed:
            raise ValueError("workbook is closed or already saved")

    @property
    def sheetnames(self):
        return [s.title for s in self.worksheets]

    @property
    def active(self):
        return self.worksheets[0] if self.worksheets else None

    def __getitem__(self, title):
        for sheet in self.worksheets:
            if sheet.title == title:
                return sheet
        raise KeyError(title)

    def create_sheet(self, title=None):
        self._ensure_open()
        if title is None:
            title = "Sheet"
            i = 1
            while title.lower() in {s.title.lower() for s in self.worksheets}:
                title = f"Sheet{i}"
                i += 1
        index = self._native.create_sheet(title)
        sheet = WriteOnlyWorksheet(self, index, title)
        self.worksheets.append(sheet)
        return sheet

    def save(self, filename, *, compression_level=4):
        self._ensure_open()
        self._native.save(fspath(filename), compression_level)
        self._closed = True
        self._native = None

    def close(self):
        self._closed = True
        self._native = None

    def __enter__(self):
        self._ensure_open()
        return self

    def __exit__(self, *args):
        self.close()

