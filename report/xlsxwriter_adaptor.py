##############################################################################
#
# Note: only required functions, class variables and properties are adapted. Others are ignored.
#
# Copyright 2020, Appcider
#
import logging
_logger = logging.getLogger(__name__)

try:
    import xlsxwriter
    import xlwt
except ImportError:
    _logger.debug('Can not import xlsxwriter and/or xlwt`.')

from xlsxwriter.workbook import Workbook
from xlsxwriter.worksheet import Worksheet
from xlsxwriter.worksheet import convert_cell_args
from xlsxwriter.compatibility import str_types
from xlsxwriter.format import Format

from .xlwt_Column import Column
from .xlwt_Row import Row
import xlwt_Style

# Functions and properties ignored:
# xs.remove_splits

class Format_adaptor(Format):
    """
    A class for writing the Excel XLSX Format file.

    """
    def __init__(self, properties=None, xf_indices=None, dxf_indices=None):
        super(Format_adaptor, self).__init__(properties, xf_indices, dxf_indices)
 
        # Note: only xlsxwriter format functions required (currently used) by Appcider are included.
        # Note: this dict is taken from xlwt_Style.py. Those keys not required by Appcider are left unchanged.
        self.xlsxwriter_format_functions = {
            'align': 'alignment', # synonym
            'alignment': {
                'dire': {
                    'general': 0,
                    'lr': 1,
                    'rl': 2,
                    },
                'direction': 'dire',
                'horz': self.set_align,
                # Possible horz Values:
                #     'general': 0,
                #     'left': 1,
                #     'center': 2,
                #     'centre': 2, # "align: horiz centre" means xf.alignment.horz is set to 2
                #     'right': 3,
                #     'filled': 4,
                #     'justified': 5,
                #     'center_across_selection': 6,
                #     'centre_across_selection': 6,
                #     'distributed': 7,
                'inde': 'IntULim(15)', # restriction: 0 <= value <= 15
                'indent': 'inde',
                'rota': [{'stacked': 255, 'none': 0, }, 'rotation_func'],
                'rotation': 'rota',
                'shri': 'bool_map',
                'shrink': 'shri',
                'shrink_to_fit': 'shri',
                'vert': self.set_align,
                # Possible vert Values:
                #     'top': 0,
                #     'center': 1,
                #     'centre': 1,
                #     'bottom': 2,
                #     'justified': 3,
                #     'distributed': 4,
                'wrap': self.set_text_wrap,
                },
            'border': 'borders',
            'borders': {
                'left':     self.set_left,
                'right':    self.set_right,
                'top':      self.set_top,
                'bottom':   self.set_bottom,
                'diag':     self.set_diag_border,
                'top_colour':       self.set_left_color,
                'bottom_colour':    self.set_right_color,
                'left_colour':      self.set_top_color,
                'right_colour':     self.set_bottom_color,
                'diag_colour':      self.set_diag_color,
                'top_color':        self.set_left_color,
                'bottom_color':     self.set_right_color,
                'left_color':       self.set_top_color,
                'right_color':      self.set_bottom_color,
                'diag_color':       self.set_diag_color,
                },
            'font': {
                'bold': self.set_bold,
                'charset': 'charset_map',
                'color':  'colour_index',
                'color_index':  'colour_index',
                'colour':  'colour_index',
                'colour_index': ['colour_map', 'colour_index_func_15'],
                'escapement': {'none': 0, 'superscript': 1, 'subscript': 2},
                'family': {'none': 0, 'roman': 1, 'swiss': 2, 'modern': 3, 'script': 4, 'decorative': 5, },
                'height': self.set_font_size, # practical limits are much narrower e.g. 160 to 1440 (8pt to 72pt)
                'italic': self.set_italic,
                'name': 'any_str_func',
                'outline': 'bool_map',
                'shadow': 'bool_map',
                'struck_out': 'bool_map',
                'underline': self.set_underline,
                },
            'pattern': {
                'pattern': self.set_pattern,
                'pattern_back_colour':  self.set_bg_color,
                'pattern_fore_colour':  self.set_fg_color,
                },
            'protection': {
                'cell_locked' :   self.set_locked,
                'formula_hidden': self.set_hidden,
                },
        }

    """
    New functions:
    """
    # Example: section=border, key=left, value=thin
    # Pre-conditions: 
    # (1) value here is already in bit value compatible with xlsxwriter (checked through xlwt_Style.py for border type and color)
    def get_xlsx_cell_format_style_func(self, section, key, value):
        # Select function from a key (string) in "xlsxwriter_format_functions" dict that matches "section" and "key" 
        func = self.xlsxwriter_format_functions[section]
        if not isinstance(func, dict):
            func = self.xlsxwriter_format_functions[func]
        func = func[key]
        func(value)

    """
    Overriden functions:
    """
    def set_font_size(self, font_size=11):
        """
        Set the Format font_size property. The default Excel font size is 11.

        Args:
            font_size: Int with font size. No default.

        Returns:
            Nothing.

        """
        self.font_size = int(font_size/18)

class Worksheet_adaptor(Worksheet):
    """
    Subclass of the XlsxWriter Worksheet class to override the default
    write() method.

    """
    def __init__(self):
        """
        Constructor.

        """
        super(Worksheet_adaptor, self).__init__()

        # Add new adaptor required class variables:
        self.parent = None
        self.panes_frozen = 0
        self.__vert_split_pos = 0 # one-indexed in xlwt
        self.__horz_split_pos = 0 # one-indexed in xlwt. i.e. self.__horz_split_pos == 1 && self.panes_frozen == 1, means 1st row is freezed
        self.Column = Column # Refers to the class. Will be instantiated as needed.
        self.Row = Row
        self.__cols = {}
        self.__rows = {}
    
    """
    Overriden functions:
    """
    @convert_cell_args
    def write(self, row, col, *args):
        # Call the parent version of write() as usual for other data.
        return super(Worksheet_adaptor, self).write(row, col, *args)

    """
    New adaptor required functions:
    """
    def set_horz_split_pos(self, value):
        self.__horz_split_pos = abs(value)
        # Freeze selected:
        if self.panes_frozen and (self.__horz_split_pos != 0 or self.__vert_split_pos != 0):
            self.freeze_panes(self.__horz_split_pos, self.__vert_split_pos)

    def set_vert_split_pos(self, value):
        self.__vert_split_pos = abs(value)
        # Unfreeze all first:
        self.freeze_panes(0, 0)
        # Freeze selected:
        if self.panes_frozen and (self.__horz_split_pos != 0 or self.__vert_split_pos != 0):
            self.freeze_panes(self.__horz_split_pos, self.__vert_split_pos)

    def col(self, indx):
        if indx not in self.__cols:
            self.__cols[indx] = self.Column(indx, self)
        return self.__cols[indx]

    def row(self, indx):
        if indx not in self.__rows:
            self.__rows[indx] = self.Row(indx, self)
        return self.__rows[indx] 

    def get_parent(self):
        return self.parent

    """
    Properties
    """
    @property
    def header_str(self):
        pass

    @header_str.setter
    def header_str(self, string_value):
        self.set_header(string_value)
    
    @property
    def footer_str(self):
        pass

    @footer_str.setter
    def footer_str(self, string_value):
        self.set_header(string_value)

class Workbook_adaptor(Workbook):
    """
    Subclass of the XlsxWriter Workbook class to override the default
    Worksheet class with our custom class.

    """

    def add_sheet(self, name=None):
        # Overwrite add_worksheet() to create a MyWorksheet object.
        worksheet = super(Workbook_adaptor, self).add_worksheet(name, Worksheet_adaptor)
        # Assign corresponding workbook of this worksheet:
        worksheet.parent = self
        return worksheet

    def easyxf(self, strg_to_parse="", num_format_str=None,
           field_sep=",", line_sep=";", intro_sep=":", esc_char="\\", debug=False):
        return xlwt_Style.easyxf(self, strg_to_parse, num_format_str, field_sep, line_sep, intro_sep, esc_char, debug)

    def add_format(self, properties=None):
        """
        Add a new Format to the Excel Workbook.
        Note: Format class used is Format_adaptor class
        Args:
            properties: The format properties.

        Returns:
            Reference to a Format object.

        """
        format_properties = self.default_format_properties.copy()

        if self.excel2003_style:
            format_properties = {'font_name': 'Arial', 'font_size': 10,
                                 'theme': 1 * -1}

        if properties:
            format_properties.update(properties)

        xf_format = Format_adaptor(format_properties,
                           self.xf_format_indices,
                           self.dxf_format_indices)

        # Store the format reference.
        self.formats.append(xf_format)

        return xf_format