# -*- coding: utf-8 -*-

# Supporting only python >= 2.6
# from __future__ import unicode_literals
# from __future__ import print_function
# from future.builtins import str as text
# from future.builtins import (range, object)

"""
:class:`ExcelTableDirective` implements the ``exceltable`` -directive.
"""
__docformat__ = 'restructuredtext'
__author__ = 'Juha Mustonen and Saptak Das'

import os
import re
import logging
from datetime import datetime

# Import required docutils modules
from docutils.parsers.rst import Directive, directives
from docutils.parsers.rst.directives.tables import ListTable
from docutils import nodes, utils, frontend
from docutils.utils import SystemMessagePropagation, Reporter

import sphinx
from sphinx.util import logging
from sphinx.application import Sphinx


# Uses Pandas (xlrd, openpyxl, odfpy, or pyxlsb) to support reading from local filesystem or URL. Pandas supports all formats below:
# * Excel 97-2003 Workbook (.xls)
# * Excel Workbook (.xlsx)
# * Excel Macro-Enabled Workbook (.xlsm)
# * Excel Workbook Template (.xltx)
# * Excel Macro-Enabled Workbook Template (.xltm)
# * Excel Binary Workbook (.xlsb)
# * OpenDocument Spreadsheet (.ods)
# * OpenDocument Text (.odt)
# * OpenDocument Formula (.odf)
import pandas as pd

basestring = (str, bytes)


def text(var):
    return var


class Messenger(Reporter):
    def __init__(self, src='sphinxcontrib.xyz'):
        settings = frontend.OptionParser().get_default_values()

        settings.report_level = 1

        Reporter.__init__(self,
                          src,
                          settings.report_level,
                          settings.halt_level,
                          stream=settings.warning_stream,
                          debug=settings.debug,
                          encoding=settings.error_encoding,
                          error_handler=settings.error_encoding_error_handler
                          )

        self.log = logging.getLogger(src)

    def debug(self, *msgs):
        # return super(Messenger, self).debug(msg)
        pass

    def info(self, *msgs):
        # return super(Messenger, self).info(msg)
        pass

    def warning(self, *msgs):
        # super(Messenger, self).warning(msg)
        return nodes.literal_block(text=self._prepare(msgs))

    def error(self, *msgs):
        # super(Messenger, self).error(msg)
        text = self._prepare(msgs)
        # self.log.error(text)
        return nodes.literal_block(text=text)

    def _prepare(self, *msgs):
        return u' '.join([text(msg) for msg in msgs])


class DirectiveTemplate(Directive):
    """
    Template intended for directive development, providing
    few handy functions
    """

    def _get_directive_path(self, path):
        """
        Returns transformed path from the directive
        option/content
        """
        source = self.state_machine.input_lines.source(
            self.lineno - self.state_machine.input_offset - 1)
        source_dir = os.path.dirname(os.path.abspath(source))
        path = os.path.normpath(os.path.join(source_dir, path))

        return utils.relative_path(None, path)


class ExcelTableDirective(ListTable, DirectiveTemplate):
    """
    ExcelTableDirective implements the directive.
    Directive allows to create RST tables from the contents
    of the Excel sheet. The functionality is very similar to
    csv-table (docutils) and xmltable (:mod:`sphinxcontrib.xmltable`).

    Example of the directive:

    .. code-block:: rest

      .. exceltable::
         :file: path/to/document.xls
         :header: 1

    """
    # required_arguments = 0
    # optional_arguments = 0
    has_content = False
    option_spec = {
        'file': directives.path,
        'selection': directives.unchanged_required,
        'header': directives.unchanged,
        'sheet': directives.unchanged,
        'class': directives.class_option,
        'widths': directives.unchanged
    }

    def run(self):
        """
        Implements the directive
        """
        # Get content and options
        file_path = self.options.get('file', None)
        selection = self.options.get('selection', ':')
        sheet = self.options.get('sheet', '0')
        header = self.options.get('header', '0')
        col_widths = self.options.get('widths', [])

        # Divide the selection into from and to values
        if u':' not in selection:
            selection += u':'
        fromcell, tocell = selection.split(u':')

        if not fromcell:
            fromcell = None

        if not tocell:
            tocell = None

        # print selection, fromcell, tocell

        if not file_path:
            return [self._report(u'file_path -option missing')]

        # Header option
        header_rows = 0
        if header and header.isdigit():
            header_rows = int(header)

        # Transform the path suitable for processing
        file_path = self._get_directive_path(file_path)
        if sheet.isdigit():
            sheet = int(sheet)

        # try:
        et = ExcelTable(file_path)
        table = et.create_table(fromcell=fromcell, tocell=tocell,
                                nheader=header_rows, widths=col_widths, 
                                sheet=sheet)
        # except Exception as e:
        # raise e.with_traceback()
        # return [msgr.error(u'Error occurred while creating table: %s' % e)]
        # pass

        # print table

        title, messages = self.make_title()
        # node = nodes.Element() # anonymous container for parsing
        # self.state.nested_parse(self.content, self.content_offset, node)

        # If empty table is created
        if not table:
            self._report('The table generated from queries is empty')
            return [nodes.paragraph(text='')]

        try:
            table_data = []

            # If there is header defined, set the header-rows param and
            # append the data in row =>. build_table_from_list handles the header generation
            if header and not header.isdigit():
                # Otherwise expect the header to be string with column names defined in
                # it, separating the values with comma
                header_rows = 1
                table_data.append([nodes.paragraph(text=hcell.strip()) for hcell in header.split(',')])

            # Put the given data in rst elements: paragraph
            for row in table['headers']:
                table_data.append([nodes.paragraph(text=cell['value']) for cell in row])

            # Iterates rows: put the given data in rst elements
            for row in table['rows']:
                row_data = []
                for cell in row:
                    class_data = ['']
                    # Node based on formatting rules
                    # NOTE: rst does not support nested, use class attribute instead

                    if cell.get('italic', False):
                        class_data.append('italic')

                    if cell.get('bold', False):
                        node = nodes.strong(text=cell['value'])
                    else:
                        node = nodes.paragraph(text=cell['value'])

                    # Add additional formatting as class attributes
                    node['classes'] = class_data
                    row_data.append([node])

                    # FIXME: style attribute does not get into writer
                    if cell.get('bgcolor', None):
                        rgb = [text(val) for val in cell['bgcolor']]
                        #node.attributes['style'] = 'background-color: rgb({});'.format(','.join(rgb))

                table_data.append(row_data)

            # If there is no data at this point, throw an error
            if not table_data:
                return [msgr.error('Selection did not return any data')]

            # Get params from data
            num_cols = len(table_data[0])

            # Get the widths for the columns:
            # 1. Use provided info, if available
            # 2. Use widths from the excelsheet
            # 3. Use default widths (equal to all)
            #
            # Get content widths from the first row of the table
            # if it fails, calculate default column widths
            if col_widths:
                col_widths = [int(width) for width in col_widths.split(',')]
            else:
                col_widths = [int(col['width']) for col in table['rows'][0]]
                col_width_total = sum(col_widths)
                col_widths = [int(width * 100 / col_width_total) for width in col_widths]

            # If still empty for some reason, use default widths
            if not col_widths:
                col_widths = self.get_column_widths(num_cols)

            stub_columns = 0

            # Sanity checks

            # Different amount of cells in first and second row (possibly header and 1 row)
            if type(header) is not int:
                if len(table_data) > 1 and len(table_data[0]) != len(table_data[1]):
                    error = msgr.error('Data amount mismatch: check the directive data and params')
                    return [error]

            self.check_table_dimensions(table_data, header_rows, stub_columns)

        except SystemMessagePropagation as detail:
            return [detail.args[0]]

        # Generate the table node from the given list of elements
        table_node = self.build_table_from_list(
            table_data, col_widths, header_rows, stub_columns)

        # Optional class parameter
        table_node['classes'] += self.options.get('class', [])

        if title:
            table_node.insert(0, title)

        # print table_node

        return [table_node] + messages


# TODO: Move away
msgr = Messenger('sphinxcontrib.exceltable')


class ExcelTable(object):
    """
    Class generates the list based table from
    the given excel-document, suitable for the directive.

    Class also implements the custom query format,
    is to use for the directive.::

      >>> import os
      >>> from sphinxcontrib import exceltable
      >>>
      >>> fo = open(os.path.join(os.path.dirname(exceltable.__file__),'../doc/example/cartoons.xls'), 'r+b')
      >>> et = exceltable.ExcelTable(fo)
      >>>
      >>> table = et.create_table(fromcell='A1', tocell='C4')
      >>> assert et.fromcell == (0, 0)
      >>> assert et.tocell == (2,3)
      >>>
      >>> table = et.create_table(fromcell='B10', tocell='B11', sheet='big')
      >>> assert et.fromcell == (1,9)
      >>> assert et.tocell == (1,10)

    """

    def __init__(self, filepath):
        """
        """
        self.filepath = filepath
        self.fromcell = (0, 0)
        self.tocell = (0, 0)
        self.df = None


    def create_table(self, fromcell=None, tocell=None, nheader=0, widths=[], sheet=0):
        """
        Creates a table (as a list) based on given query and columns

        fromcell:
          The index of the cell where to begin. The default
          is from the beginning of the data set (0, 0).

        tocell:
          The index of the cell where to end. The default
          is to the end of the data set.

        nheader:
          Number of lines which are considered as a header lines.
          Normally, the value is 0 (default) or 1.

        widths:
          List of widths for the columns. The default is to use
          equal widths for all columns.

        sheet:
          Name or index of the sheet as string/unicode. The index starts from the 0
          and is the default value.

          et.create_table(fromcell='A1', tocell='C4', nheader=1, widths=[40, 30, 30], sheet='Sheet1', date_format='%Y-%m-%d')
        """
        rows = []

        # Name selection, like: 'A1' or 'AB12'
        if isinstance(fromcell, basestring):
            match = re.match(r'(?P<chars>[A-Z]+)(?P<nums>[1-9]+[0-9]*)', fromcell)
            if match:
                parts = (match.group('chars'), int(match.group('nums')))
                fromcell = toindex(*parts)
            else:
                fromcell = tuple([int(num) for num in fromcell.split(u',')])

        # Name selection, like: 'A1' or 'AB12'
        if isinstance(tocell, basestring):
            match = re.match(r'(?P<chars>[A-Z]+)(?P<nums>[1-9]+[0-9]*)', tocell)
            if match:
                parts = (match.group('chars'), int(match.group('nums')))
                tocell = toindex(*parts)
            else:
                tocell = tuple([int(num) for num in tocell.split(u',')])

        usecols = list(range(fromcell[0], tocell[0] + 1)) if fromcell and tocell else None
        skiprows = fromcell[1] if fromcell else None
        self.df = pd.read_excel(self.filepath, sheet_name=sheet, header=None, index_col=None, usecols=usecols, skiprows=skiprows)

        # Choose the first column based on fromcell if usecols not used.
        if not usecols and fromcell:
            self.df = self.df.iloc[:, fromcell[0]:]

        # Relabel columns to 0, 1, 2, ...
        self.df.columns = list(range(len(self.df.columns)))

        # Cut the df to the correct size
        if fromcell and tocell:
            self.df = self.df.iloc[:tocell[1] - fromcell[1] + 1, :]

        # Update fromcell and tocell if not given
        if not fromcell:
            fromcell = (0, 0)
        if not tocell:
            tocell = (fromcell[0] + len(self.df.columns) - 1, fromcell[1] + len(self.df.index) - 1)

        # Iterate columns
        rows = {'headers': [], 'rows': []}

        for row_num in range(fromcell[1], tocell[1] + 1):

            # Iterate rows within column
            cols = []
            for cnum in range(fromcell[0], tocell[0] + 1):
                # Value will always be determined from pandas.
                value = self.df.iloc[row_num - fromcell[1], cnum - fromcell[0]]
                width = widths[cnum - fromcell[0]] if len(widths) == tocell[0] + 1 - fromcell[0] else 0
                if width == 0:
                    width = 20 # Default width if not specified
                cell_data = {'type': 'row', 'width': width, 'value': value}

                # If header row
                if row_num < nheader:
                    cell_data['type'] = 'header'

                # Get more format info for the cell
                # TODO: Can add formatting for specific file types using various engines later.
                # cell_data.update(self._get_formatting(cell))

                cols.append(cell_data)

            # The first column is assumed to be all headers.
            if cols[0]['type'] == 'header':
                # Make columns bolded.
                for col in cols:
                    col['bold'] = True
                rows['headers'].append(cols)
            else:
                rows['rows'].append(cols)

        # widths_together = sum([cell['width'] for cols in rows])
        # print widths_together
        # widths = [round(val * 100.0 / widths_together) for val in widths]

        # Store into object for validation purposes
        self.fromcell = fromcell
        self.tocell = tocell
        return rows


def toindex(col, row):
    """
    Calculates the index number from
    the Excel column name. Examples:

      >>> from sphinxcontrib import exceltable
      >>> exceltable.toindex('A', 1)
      (0, 0)
      >>> exceltable.toindex('B', 10)
      (1, 9)
      >>> exceltable.toindex('Z', 2)
      (25, 1)
      >>> exceltable.toindex('AA', 27)
      (26, 26)
      >>> exceltable.toindex('AB', 1)
      (27, 0)

    .. NOTE::

       Following the naming in Excel/OOCalc,
       the row 'index' starts from the 1 and not from 0

    """
    a2z = 'ABCDEFGHIJLKMNOPQRSTUVWXYZ'

    total = 0
    mult = 0
    for char in col:
        total += (a2z.find(char) + (26 * mult))
        mult += 1

    return total, row - 1


def toname(colx, rowy):
    """
    Opposite to `toindex`
    """
    # Convert int colx into a Excel column name
    # e.g. 0 -> A, 1 -> B, ..., 25 -> Z, 26 -> AA, 27 -> AB
    col_name = ''
    while colx >= 0:
        col_name = chr(ord('A') + colx % 26) + col_name
        colx = int(colx / 26) - 1
    return col_name, rowy + 1


def setup(app: Sphinx):
    """
    Extension setup, called by Sphinx
    """

    # Sphinx 0.5 support
    if sphinx.__version__.startswith('0.5'):
        app.add_directive('exceltable', ExcelTableDirective, 0, (0, 0, 0))
    else:
        app.add_directive('exceltable', ExcelTableDirective)
