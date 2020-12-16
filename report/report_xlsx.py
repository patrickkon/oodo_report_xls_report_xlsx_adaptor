# -*- coding: utf-8 -*-
# Copyright 2015 ACSONE SA/NV (<http://acsone.eu>)
# License AGPL-3.0 or later (http://www.gnu.org/licenses/agpl.html).

from cStringIO import StringIO
from datetime import datetime
from openerp.osv.fields import datetime as datetime_field
from openerp.tools import DEFAULT_SERVER_DATETIME_FORMAT
from openerp.report.report_sxw import report_sxw
from openerp.api import Environment
from openerp import pooler

import logging
_logger = logging.getLogger(__name__)

try:
    import xlsxwriter
except ImportError:
    _logger.debug('Can not import xlsxwriter`.')

from .xlsxwriter_adaptor import Workbook_adaptor as Workbook

xls_types_default = {
    'bool': False,
    'date': None,
    'text': '',
    'number': 0,
}

class AttrDict(dict):
    def __init__(self, *args, **kwargs):
        super(AttrDict, self).__init__(*args, **kwargs)
        self.__dict__ = self

class ReportXlsx(report_sxw):
    # header/footer
    hf_params = {
        'font_size': 8,
        'font_style': 'I',  # B: Bold, I:  Italic, U: Underline
    }

    # styles
    _pfc = 'light_yellow'  # default pattern fore_color
    _bc = 'gray25'   # borders color
    decimal_format = '#,##0.00'
    date_format = 'YYYY-MM-DD'
    xls_styles = {
        'xls_title': 'font: bold true, height 240;',
        'bold': 'font: bold true;',
        'underline': 'font: underline true;',
        'italic': 'font: italic true;',
        'fill': 'pattern: pattern solid, fore_color %s;' % _pfc,
        'fill_blue': 'pattern: pattern solid, fore_color light_blue;',
        'fill_grey': 'pattern: pattern solid, fore_color gray25;',
        'borders_all':
            'borders: '
            'left thin, right thin, top thin, bottom thin, '
            'left_colour %s, right_colour %s, '
            'top_colour %s, bottom_colour %s;'
            % (_bc, _bc, _bc, _bc),
        'left': 'align: horz left;',
        'center': 'align: horz center;',
        'right': 'align: horz right;',
        'wrap': 'align: wrap true;',
        'top': 'align: vert top;',
        'bottom': 'align: vert bottom;',
    }

    def create(self, cr, uid, ids, data, context=None):
        self.env = Environment(cr, uid, context)
        self.pool = pooler.get_pool(cr.dbname)
        self.cr = cr
        self.uid = uid
        self.context = context
        report_obj = self.env['ir.actions.report.xml']
        report = report_obj.search([('report_name', '=', self.name[7:])])
        if report.ids:
            self.title = report.name
            if report.report_type == 'xls':
                return self.create_xlsx_report(ids, data, report)
        elif context.get('xls_export'):
            # use model from 'data' when no ir.actions.report.xml entry
            self.table = data.get('model') or self.table
            _logger.error("in context xlsx_export == 1 with self.table %s", str(self.table))
            return self.create_xlsx_report(ids, data, context)
        return super(ReportXlsx, self).create(cr, uid, ids, data, context)

    def create_xlsx_report(self, ids, data, report):
        _logger.error("in xlsx report with env: %s", str(self.env.cr))
        self.parser_instance = self.parser(
            self.env.cr, self.env.uid, self.name2, self.env.context)
        objs = self.getObjects(
            self.env.cr, self.env.uid, ids, self.env.context)
        self.parser_instance.set_context(objs, data, ids, 'xlsx')
        _p = AttrDict(self.parser_instance.localcontext)
        objs = self.parser_instance.localcontext['objects']
        file_data = StringIO()
        workbook = Workbook(file_data)
        _xs = self.xls_styles
        self.xls_headers = {
            'standard': '',
        }
        report_date = datetime_field.context_timestamp(
            self.cr, self.uid, datetime.now(), self.context)
        report_date = report_date.strftime(DEFAULT_SERVER_DATETIME_FORMAT)
        self.xls_footers = {
            'standard': report_date
        }
        self.generate_xls_report(_p, _xs, data, objs, workbook)
        workbook.close()
        file_data.seek(0)
        
        return (file_data.read(), 'xlsx')

    # Commonly, _p, objects are left blank:
    # _xs: style sheet dict
    # data: context data
    # wb: workbook instance
    def generate_xls_report(self, _p, _xs, data, objects, wb):
        raise NotImplementedError()
    
    # def generate_xlsx_report(self, workbook, data, objs):
    #     raise NotImplementedError()

    # Note: column formula feature is omitted
    def xls_row_template(self, specs, wanted_list):
        """
        Returns a row template.

        Input :
        - 'wanted_list': list of Columns that will be returned in the
                         row_template
        - 'specs': list with Column Characteristics
            0: Column Name (from wanted_list)
            1: Column Colspan
            2: Column Size (unit = the width of the character ’0′
                            as it appears in the sheet’s default font)
            3: Column Type
            4: Column Data
            5: Column Formula (or 'None' for Data)
            6: Column Style
        """
        r = []
        col = 0
        for w in wanted_list:
            found = False
            for s in specs:
                if s[0] == w:
                    found = True
                    s_len = len(s)
                    # Set cell data:
                    size = s[1]
                    col_width = s[2]
                    datatype = s[3]
                    data = s[4]
                    # Set custom cell style
                    if s_len > 6 and s[6] is not None:
                        style = s[6]
                    else:
                        style = None
                    # starting column, colspan, column type, column data, column style
                    r.append((col, size, col_width, datatype, data, style))
                    col += s[1]
                    break
            if not found:
                _logger.warn("report_xls.xls_row_template, "
                             "column '%s' not found in specs", w)
        return r
    
    # Note: column formula feature is omitted
    def xls_write_row(self, ws, row_pos, row_data,
                      row_style=None, set_column_size=None):
        for col, size, col_width, datatype, data, style in row_data:
            style = style or row_style
            if not data:
                # if no data, use default values
                data = xls_types_default[datatype]
            if size != 1:
                if style:
                    ws.merge_range(row_pos, col, row_pos, col + size - 1, data, style)
                else:
                    ws.merge_range(row_pos, col, row_pos, col + size - 1, data)
            else:
                if style:
                    ws.write(row_pos, col, data, style)
                else:
                    ws.write(row_pos, col, data)
            if set_column_size:
                ws.col(col).width = col_width * 256
        return row_pos + 1
