# -*- coding: utf-8 -*-

import datetime as dt
from decimal import Decimal

import six

class Row(object):
    __slots__ = [# private variables
                 "__idx",
                 "__parent",
                 "__parent_wb",
                 "__cells",
                 "__min_col_idx",
                 "__max_col_idx",
                 "__xf_index",
                 "__has_default_xf_index",
                 "__height_in_pixels",
                 "__height",
                 # public variables
                 "height",
                 "has_default_height",
                 "height_mismatch",
                 "level",
                 "collapse",
                 "hidden",
                 "space_above",
                 "space_below"]

    def __init__(self, rowx, parent_sheet):
        if not (isinstance(rowx, six.integer_types) and 0 <= rowx <= 1048576):
            raise ValueError("row index was %r, not allowed by .xlsx format" % rowx)
        self.__idx = rowx # zero-indexed in xlwt
        self.__parent = parent_sheet
        self.__parent_wb = parent_sheet.get_parent()
        self.__cells = {}
        self.__min_col_idx = 0
        self.__max_col_idx = 0
        self.__xf_index = 0x0F
        self.__has_default_xf_index = 0
        self.__height_in_pixels = 0x11

        self.height = 0x00FF
        self.__height = 0x00FF
        self.has_default_height = 0x00
        self.height_mismatch = 0
        self.level = 0
        self.collapse = 0
        self.hidden = 0
        self.space_above = 0
        self.space_below = 0

    def __adjust_height(self, style):
        twips = style.font.height
        points = float(twips)/20.0
        # Cell height in pixels can be calcuted by following approx. formula:
        # cell height in pixels = font height in points * 83/50 + 2/5
        # It works when screen resolution is 96 dpi
        pix = int(round(points*83.0/50.0 + 2.0/5.0))
        if pix > self.__height_in_pixels:
            self.__height_in_pixels = pix
    
    def set_height(self, height):
        self.__height = int(height/16) 
        self.__parent.set_row(self.__idx, self.__height)
    
    def get_height(self):
        return self.__height

    height = property(get_height, set_height)

    def __adjust_bound_col_idx(self, *args):
        for arg in args:
            iarg = int(arg)
            if not ((0 <= iarg <= 255) and arg == iarg):
                raise ValueError("column index (%r) not an int in range(256)" % arg)
            sheet = self.__parent
            if iarg < self.__min_col_idx:
                self.__min_col_idx = iarg
            if iarg > self.__max_col_idx:
                self.__max_col_idx = iarg
            if iarg < sheet.first_used_col:
                sheet.first_used_col = iarg
            if iarg > sheet.last_used_col:
                sheet.last_used_col = iarg

    def __excel_date_dt(self, date):
        adj = False
        if isinstance(date, dt.date):
            if self.__parent_wb.dates_1904:
                epoch_tuple = (1904, 1, 1)
            else:
                epoch_tuple = (1899, 12, 31)
                adj = True
            if isinstance(date, dt.datetime):
                epoch = dt.datetime(*epoch_tuple)
            else:
                epoch = dt.date(*epoch_tuple)
        else: # it's a datetime.time instance
            date = dt.datetime.combine(dt.datetime(1900, 1, 1), date)
            epoch = dt.datetime(1900, 1, 1)
        delta = date - epoch
        xldate = delta.days + delta.seconds / 86400.0
        # Add a day for Excel's missing leap day in 1900
        if adj and xldate > 59:
            xldate += 1
        return xldate

    def get_height_in_pixels(self):
        return self.__height_in_pixels


    def set_style(self, style):
        self.__adjust_height(style)
        self.__xf_index = self.__parent_wb.add_style(style)
        self.__has_default_xf_index = 1


    def get_xf_index(self):
        return self.__xf_index


    def get_cells_count(self):
        return len(self.__cells)


    def get_min_col(self):
        return self.__min_col_idx


    def get_max_col(self):
        return self.__max_col_idx

    def insert_cell(self, col_index, cell_obj):
        if col_index in self.__cells:
            if not self.__parent._cell_overwrite_ok:
                msg = "Attempt to overwrite cell: sheetname=%r rowx=%d colx=%d" \
                    % (self.__parent.name, self.__idx, col_index)
                raise Exception(msg)
            prev_cell_obj = self.__cells[col_index]
            sst_idx = getattr(prev_cell_obj, 'sst_idx', None)
            if sst_idx is not None:
                self.__parent_wb.del_str(sst_idx)
        self.__cells[col_index] = cell_obj

    def insert_mulcells(self, colx1, colx2, cell_obj):
        self.insert_cell(colx1, cell_obj)
        for col_index in range(colx1+1, colx2+1):
            self.insert_cell(col_index, None)

    def get_index(self):
        return self.__idx