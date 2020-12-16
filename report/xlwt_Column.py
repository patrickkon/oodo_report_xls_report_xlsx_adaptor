# -*- coding: utf-8 -*-

# Note a Column class instance is instantiated every single time ws.col(index) is called, where index has not been used before. This is seen in  xlwt module -> Worksheet.py -> def col(self, indx)
class Column(object):
    # TODO: change or remove _parent_wb since get_parent is an invalid function in xlxswriter package.
    def __init__(self, colx, parent_sheet):
        if not(isinstance(colx, int) and 0 <= colx <= 16384):
            raise ValueError("column index (%r) not an int in range(16384)" % colx)
        self._index = colx # zero-indexed in xlwt
        self._parent = parent_sheet
        self._parent_wb = parent_sheet.get_parent()
        self._xf_index = 0x0F

        self.width = 0x0B92
        self.__width = 0x0B92
        self.hidden = 0
        self.level = 0
        self.collapse = 0
        self.user_set = 0
        self.best_fit = 0
        self.unused = 0

    def set_width(self, width):
        if not(isinstance(width, int) and 0 <= width <= 65535):
            raise ValueError("column width (%r) not an int in range(65536)" % width)
        self._width = int(width/300)
        # Override with xlsxwriter functionality: 
        # In xlwt, column width is adjusted over a single column only.
        self._parent.set_column(self._index, self._index, self._width)

    def get_width(self):
        return self._width

    width = property(get_width, set_width)

    def set_style(self, style):
        self._xf_index = self._parent_wb.add_style(style)

    def width_in_pixels(self):
        # *** Approximation ****
        return int(round(self.width * 0.0272 + 0.446, 0))

