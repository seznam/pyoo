
import contextlib
import datetime
import unittest2

import pyoo


desktop = None


def setUpModule():
    global desktop
    desktop = pyoo.Desktop()


class BaseTestCase(unittest2.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.document = desktop.create_spreadsheet()

    @classmethod
    def tearDownClass(cls):
        cls.document.close()


class CellRangeTestCase(BaseTestCase):

    def setUp(self):
        self.sheet = self.document.sheets[0]

    # Test conversion to string

    def test_cell_unicode(self):
        cell = self.sheet[0,0]
        self.assertEqual(u'$A$1', unicode(cell))
        self.assertEqual(u'$A$1', unicode(cell.address))

    def test_cell_str(self):
        cell = self.sheet[0,0]
        self.assertEqual('$A$1', str(cell))
        self.assertEqual('$A$1', str(cell.address))

    def test_cell_repr(self):
        cell = self.sheet[0,0]
        self.assertEqual("<Cell: '$A$1'>", repr(cell))
        self.assertEqual("<SheetAddress: '$A$1'>", repr(cell.address))

    # Test cell range slicing and indexing:

    def test_row_index_must_be_int(self):
        with self.assertRaises(TypeError):
            self.sheet['x',:]

    def test_col_index_must_be_int(self):
        with self.assertRaises(TypeError):
            self.sheet[:,'x']

    def test_row_start_index_must_be_int(self):
        with self.assertRaises(TypeError):
            self.sheet['x':,:]

    def test_row_end_index_must_be_int(self):
        with self.assertRaises(TypeError):
            self.sheet[:'x',:]

    def test_col_start_index_must_be_int(self):
        with self.assertRaises(TypeError):
            self.sheet['x':,:]

    def test_col_end_index_must_be_int(self):
        with self.assertRaises(TypeError):
            self.sheet[:'x',:]

    def test_row_step_is_not_supported(self):
        with self.assertRaises(NotImplementedError):
            self.sheet[::1,:]

    def test_col_step_is_not_supported(self):
        with self.assertRaises(NotImplementedError):
            self.sheet[:,::1]

    def test_row_index_must_be_lt_row_count(self):
        with self.assertRaises(IndexError):
            self.sheet[1048576,:]

    def test_col_index_must_be_lt_col_count(self):
        with self.assertRaises(IndexError):
            self.sheet[:,1024]

    def test_two_dimensions_only(self):
        with self.assertRaises(ValueError):
            self.sheet[:,:,:]

    def test_get_tabular_range_from_sheet(self):
        cells = self.sheet[10:20,1:6]
        self.assertEqual('$B$11:$F$20', str(cells.address))

    def test_get_default_tabular_range_from_sheet(self):
        cells = self.sheet[:,:]
        self.assertEqual('$A$1:$AMJ$1048576', str(cells.address))

    def test_get_negative_tabular_range_from_sheet(self):
        cells = self.sheet[-20:-10,-6:-1]
        self.assertEqual('$AME$1048557:$AMI$1048566', str(cells.address))

    def test_get_tabular_range_from_tabular_range(self):
        cells1 = self.sheet[10:20,1:6]
        cells2 = cells1[1:9,1:4]
        self.assertEqual('$C$12:$E$19', str(cells2.address))

    def test_get_horizontal_range_from_sheet(self):
        row = self.sheet[10,1:6]
        self.assertEqual('$B$11:$F$11', str(row.address))

    def test_get_default_horizontal_range_from_sheet(self):
        row = self.sheet[10,:]
        self.assertEqual('$A$11:$AMJ$11', str(row.address))

    def test_get_negative_horizontal_range_from_sheet(self):
        row = self.sheet[-10,-6:-1]
        self.assertEqual('$AME$1048567:$AMI$1048567', str(row.address))

    def test_get_horizontal_range_from_horizontal_range(self):
        row1 = self.sheet[10,1:6]
        row2 = row1[1:4]
        self.assertEqual('$C$11:$E$11', str(row2.address))

    def test_get_default_horizontal_range_from_horizontal_range(self):
        row1 = self.sheet[10,1:6]
        row2 = row1[:]
        self.assertEqual('$B$11:$F$11', str(row2.address))

    def test_get_negative_horizontal_range_from_horizontal_range(self):
        row1 = self.sheet[10,1:6]
        row2 = row1[-4:-1]
        self.assertEqual('$C$11:$E$11', str(row2.address))

    def test_get_vertical_range_from_sheet(self):
        column = self.sheet[10:20,1]
        self.assertEqual('$B$11:$B$20', str(column.address))

    def test_get_default_vertical_range_from_sheet(self):
        column = self.sheet[:,1]
        self.assertEqual('$B$1:$B$1048576', str(column.address))

    def test_get_negative_vertical_range_from_sheet(self):
        column = self.sheet[-20:-10,-1]
        self.assertEqual('$AMJ$1048557:$AMJ$1048566', str(column.address))

    def test_get_vertical_range_from_vertical_range(self):
        column1 = self.sheet[10:20,1]
        column2 = column1[1:9]
        self.assertEqual('$B$12:$B$19', str(column2.address))

    def test_get_default_vertical_range_from_vertical_range(self):
        column1 = self.sheet[10:20,1]
        column2 = column1[:]
        self.assertEqual('$B$11:$B$20', str(column2.address))

    def test_get_negative_vertical_range_from_vertical_range(self):
        column1 = self.sheet[10:20,1]
        column2 = column1[-9:-1]
        self.assertEqual('$B$12:$B$19', str(column2.address))

    def test_get_cell_from_sheet(self):
        cell = self.sheet[10,1]
        self.assertEqual('$B$11', str(cell.address))

    def test_get_negative_cell_from_sheet(self):
        cell = self.sheet[-10,-1]
        self.assertEqual('$AMJ$1048567', str(cell.address))

    def test_get_cell_from_range(self):
        cells = self.sheet[10:20,1:6]
        cell = cells[1,1]
        self.assertEqual('$C$12', str(cell.address))

    def test_get_negative_cell_from_range(self):
        cells = self.sheet[10:20,1:6]
        cell = cells[-1,-1]
        self.assertEqual('$F$20', str(cell.address))

    def test_get_cell_from_horizontal_range(self):
        row = self.sheet[10,:]
        cell = row[1]
        self.assertEqual('$B$11', str(cell.address))

    def test_get_negative_cell_from_horizontal_range(self):
        row = self.sheet[10,:]
        cell = row[-1]
        self.assertEqual('$AMJ$11', str(cell.address))

    def test_get_cell_from_vertical_range(self):
        column = self.sheet[:,1]
        cell = column[10]
        self.assertEqual('$B$11', str(cell.address))

    def test_get_negative_cell_from_vertical_range(self):
        column = self.sheet[:,1]
        cell = column[-10]
        self.assertEqual('$B$1048567', str(cell.address))

    def test_get_simple_tabular_range_from_sheet(self):
        cells = self.sheet[10:20]
        self.assertEqual('$A$11:$AMJ$20', str(cells.address))

    def test_get_simple_horizontal_range_from_sheet(self):
        row = self.sheet[10]
        self.assertEqual('$A$11:$AMJ$11', str(row.address))

    def test_get_simple_tabular_range_from_tabular_range(self):
        cells1 = self.sheet[10:20,1:6]
        cells2 = cells1[1:9]
        self.assertEqual('$B$12:$F$19', str(cells2.address))

    def test_get_simple_horizontal_range_from_tabular_range(self):
        cells = self.sheet[10:20,1:6]
        row = cells[0]
        self.assertEqual('$B$11:$F$11', str(row.address))

    def test_tabular_range_len(self):
        cells = self.sheet[10:20,1:6]
        self.assertEqual(10, len(cells))

    def test_horizontal_range_len(self):
        row = self.sheet[10,1:6]
        self.assertEqual(5, len(row))

    def test_vertical_range_len(self):
        column = self.sheet[10:20,1]
        self.assertEqual(10, len(column))

    # Test cursor movement (internal API is used here)

    def test_cursor_resize_move(self):
        addr = self.sheet._get_target().RangeAddress
        self.assertEqual(0, addr.StartRow)
        self.assertEqual(0, addr.StartColumn)
        self.assertEqual(0xfffff, addr.EndRow)
        self.assertEqual(0x3ff, addr.EndColumn)
        addr = self.sheet[-1,-1]._get_target().RangeAddress
        self.assertEqual(0xfffff, addr.StartRow)
        self.assertEqual(0x3ff, addr.StartColumn)
        self.assertEqual(0xfffff, addr.EndRow)
        self.assertEqual(0x3ff, addr.EndColumn)

    def test_cursor_move_resize(self):
        addr = self.sheet[-1,-1]._get_target().RangeAddress
        self.assertEqual(0xfffff, addr.StartRow)
        self.assertEqual(0x3ff, addr.StartColumn)
        self.assertEqual(0xfffff, addr.EndRow)
        self.assertEqual(0x3ff, addr.EndColumn)
        addr = self.sheet._get_target().RangeAddress
        self.assertEqual(0, addr.StartRow)
        self.assertEqual(0, addr.StartColumn)
        self.assertEqual(0xfffff, addr.EndRow)
        self.assertEqual(0x3ff, addr.EndColumn)

    def test_cursor_resize_move_resize(self):
        addr = self.sheet[-1,:]._get_target().RangeAddress
        self.assertEqual(0xfffff, addr.StartRow)
        self.assertEqual(0, addr.StartColumn)
        self.assertEqual(0xfffff, addr.EndRow)
        self.assertEqual(0x3ff, addr.EndColumn)
        addr = self.sheet[:,-1]._get_target().RangeAddress
        self.assertEqual(0, addr.StartRow)
        self.assertEqual(0x3ff, addr.StartColumn)
        self.assertEqual(0xfffff, addr.EndRow)
        self.assertEqual(0x3ff, addr.EndColumn)


    # Test different cell types

    def test_invalid_cell_value(self):
        cell = self.sheet[0, 0]
        with self.assertRaises(ValueError):
            cell.value = object()

    def test_empty_cell_value(self):
        cell = self.sheet[0, 0]
        cell.value = None
        self.assertEqual(None, cell.value)
        self.assertEqual('#N/A', cell.formula)
        self.assertEqual(None, cell.date)
        self.assertEqual(None, cell.time)

    def test_number_cell_value(self):
        cell = self.sheet[0, 0]
        cell.value = 1
        self.assertEqual(1, cell.value)
        self.assertEqual('1', cell.formula)

    def test_text_cell_value(self):
        cell = self.sheet[0, 0]
        cell.value = 'hello'
        self.assertEqual('hello', cell.value)
        self.assertEqual('hello', cell.formula)
        self.assertEqual(None, cell.date)
        self.assertEqual(None, cell.time)

    def test_datetime_cell_value(self):
        cell = self.sheet[0, 0]
        cell.value = datetime.datetime(1985, 5, 6, 23, 55)
        # Should be almost equal because dates and times are represented as floats
        self.assertEqual(datetime.datetime(1985, 5, 6, 23, 55), cell.date.replace(microsecond=0))

    def test_date_cell_value(self):
        cell = self.sheet[0, 0]
        cell.value = datetime.date(1985, 5, 6)
        # Should be almost equal because dates and times are represented as floats
        self.assertEqual(datetime.datetime(1985, 5, 6, 0, 0, 0), cell.date.replace(microsecond=0))

    def test_time_cell_value(self):
        cell = self.sheet[0, 0]
        cell.value = datetime.time(23, 55, 01)
        # Should be almost equal because dates and times are represented as floats
        self.assertEqual(datetime.time(23, 55, 01), cell.time.replace(microsecond=0))

    def test_invalid_cell_formula(self):
        cell = self.sheet[0, 0]
        with self.assertRaises(ValueError):
            cell.formula = object()

    def test_empty_cell_formula(self):
        cell = self.sheet[0, 0]
        cell.formula = None
        self.assertEqual('', cell.value)
        self.assertEqual('', cell.formula)

    def test_number_cell_formula(self):
        cell = self.sheet[0, 0]
        cell.formula = 1
        self.assertEqual(1, cell.value)
        self.assertEqual('1', cell.formula)

    def test_text_cell_formula(self):
        cell = self.sheet[0, 0]
        cell.formula = 'hello'
        self.assertEqual('hello', cell.value)
        self.assertEqual('hello', cell.formula)

    def test_formula_cell_formula(self):
        cell = self.sheet[0, 0]
        cell.formula = '=1'
        self.assertEqual(1, cell.value)
        self.assertEqual('=1', cell.formula)

    def test_datetime_cell_formula(self):
        cell = self.sheet[0, 0]
        cell.formula = datetime.datetime(1985, 5, 6, 23, 55)
        # Should be almost equal because dates and times are represented as floats
        self.assertEqual(datetime.datetime(1985, 5, 6, 23, 55), cell.date.replace(microsecond=0))

    def test_date_cell_formula(self):
        cell = self.sheet[0, 0]
        cell.formula = datetime.date(1985, 5, 6)
        # Should be almost equal because dates and times are represented as floats
        self.assertEqual(datetime.datetime(1985, 5, 6, 0, 0, 0), cell.date.replace(microsecond=0))

    def test_time_cell_formula(self):
        cell = self.sheet[0, 0]
        cell.formula = datetime.time(23, 55, 01)
        # Should be almost equal because dates and times are represented as floats
        self.assertEqual(datetime.time(23, 55, 01), cell.time.replace(microsecond=0))

    # Test data access:

    def test_tabular_range_data(self):
        data = [[i * j for j in xrange(5)] for i in xrange(10)]
        self.sheet[10:20,1:6].values = data
        self.assertEqual(36, self.sheet[19,5].value)
        self.assertEqual(36, self.sheet[10:20,1:6].values[-1][-1])

    def test_horizontal_range_data(self):
        data = [i + 100 for i in xrange(5)]
        self.sheet[10,1:6].values = data
        self.assertEqual(104, self.sheet[10,5].value)
        self.assertEqual(104, self.sheet[10,1:6].values[-1])

    def test_vertical_range_data(self):
        data = [i + 200 for i in xrange(10)]
        self.sheet[10:20,1].values = data
        self.assertEqual(209, self.sheet[19,1].value)
        self.assertEqual(209, self.sheet[10:20,1].values[-1])

    def test_cell_data(self):
        self.sheet[10,1].value = 300
        self.assertEqual(300, self.sheet[10,1].value)

    def test_tabular_range_formulas(self):
        formulas = [['=%d*%d' % (i, j) for j in xrange(5)] for i in xrange(10)]
        self.sheet[10:20,1:6].formulas = formulas
        self.assertEqual('=9*4', self.sheet[19,5].formula)
        self.assertEqual('=9*4', self.sheet[10:20,1:6].formulas[-1][-1])

    def test_horizontal_range_formulas(self):
        formulas = ['=%d+100' % i for i in xrange(5)]
        self.sheet[10,1:6].formulas = formulas
        self.assertEqual('=4+100', self.sheet[10,5].formula)
        self.assertEqual('=4+100', self.sheet[10,1:6].formulas[-1])

    def test_vertical_range_formulas(self):
        formulas = ['=%d+200' % i for i in xrange(10)]
        self.sheet[10:20,1].formulas = formulas
        self.assertEqual('=9+200', self.sheet[19,1].formula)
        self.assertEqual('=9+200', self.sheet[10:20,1].formulas[-1])

    def test_cell_formulas(self):
        self.sheet[10,1].formula = '=300'
        self.assertEqual('=300', self.sheet[10,1].formula)

    # Test cell formatting and manipulation:

    def test_cell_position(self):
        position = self.sheet[0,0].position
        self.assertEqual(0, position.x)
        self.assertEqual(0, position.y)
        self.assertTrue(0 < position.width < 10000)
        self.assertTrue(0 < position.height < 10000)

    def test_cells_merging(self):
        self.assertFalse(self.sheet[0:2,0:2].is_merged)
        self.sheet[0:2,0:2].is_merged = True
        self.assertTrue(self.sheet[0:2,0:2].is_merged)
        self.sheet[0:2,0:2].is_merged = False
        self.assertFalse(self.sheet[0:2,0:2].is_merged)

    def test_text_align(self):
        cells = self.sheet[0:2,0:2]
        self.assertEqual(pyoo.TEXT_ALIGN_STANDARD, cells.text_align)
        cells.text_align = pyoo.TEXT_ALIGN_CENTER
        self.assertEqual(pyoo.TEXT_ALIGN_CENTER, cells.text_align)
        cells.text_align = pyoo.TEXT_ALIGN_STANDARD
        self.assertEqual(pyoo.TEXT_ALIGN_STANDARD, cells.text_align)

    def test_font_size(self):
        cells = self.sheet[0:2,0:2]
        self.assertEqual(10, cells.font_size)
        self.sheet.font_size = 12
        self.assertEqual(12, cells.font_size)
        self.sheet.font_size = 10
        self.assertEqual(10, cells.font_size)

    def test_font_weight(self):
        cells = self.sheet[0:2,0:2]
        self.assertEqual(pyoo.FONT_WEIGHT_NORMAL, cells.font_weight)
        self.sheet.font_weight = pyoo.FONT_WEIGHT_BLACK
        self.assertEqual(pyoo.FONT_WEIGHT_BLACK, cells.font_weight)
        self.sheet.font_weight = pyoo.FONT_WEIGHT_NORMAL
        self.assertEqual(pyoo.FONT_WEIGHT_NORMAL, cells.font_weight)

    def test_underline(self):
        cells = self.sheet[0:2,0:2]
        self.assertEqual(pyoo.UNDERLINE_NONE, cells.underline)
        self.sheet.underline = pyoo.UNDERLINE_DOUBLE
        self.assertEqual(pyoo.UNDERLINE_DOUBLE, cells.underline)
        self.sheet.underline = pyoo.UNDERLINE_NONE
        self.assertEqual(pyoo.UNDERLINE_NONE, cells.underline)

    def test_text_color(self):
        cells = self.sheet[0:2,0:2]
        self.assertTrue(cells.text_color is None)
        self.sheet.text_color = 0xff0000
        self.assertEqual(0xff0000, cells.text_color)
        self.sheet.text_color = None
        self.assertTrue(cells.text_color is None)

    def test_background_color(self):
        cells = self.sheet[0:2,0:2]
        self.assertTrue(cells.background_color is None)
        self.sheet.background_color = 0xff0000
        self.assertEqual(0xff0000, cells.background_color)
        self.sheet.background_color = None
        self.assertTrue(cells.background_color is None)

    def test_border_width(self):
        cells = self.sheet[10:20,10:20]
        cells.border_width = 100
        # Border widths are approximate
        self.assertAlmostEqual(100, cells.border_width, delta=10)
        self.assertAlmostEqual(100, cells.inner_border_width, delta=10)
        self.assertAlmostEqual(100, cells[0,0].border_width, delta=10)
        cells.inner_border_width = 200
        self.assertAlmostEqual(200, cells.inner_border_width, delta=10)
        self.assertEqual(0, cells.border_width)
        self.assertEqual(0, cells.border_width)
        self.assertEqual(0, cells[0,0].border_width)

    # Test number formats:

    def test_int_format(self):
        fmt = self.document.get_locale().format(pyoo.FORMAT_INT)
        self.sheet[0:2,0:2].number_format = fmt
        self.assertEqual(fmt, self.sheet[0:2,0:2].number_format)
        self.assertEqual(fmt, self.sheet[0,0].number_format)

    def test_float_format(self):
        fmt = self.document.get_locale().format(pyoo.FORMAT_FLOAT)
        self.sheet[0:2,0:2].number_format = fmt
        self.assertEqual(fmt, self.sheet[0:2,0:2].number_format)
        self.assertEqual(fmt, self.sheet[0,0].number_format)

    def test_date_format(self):
        fmt = self.document.get_locale().format(pyoo.FORMAT_DATE)
        self.sheet[0:2,0:2].number_format = fmt
        self.assertEqual(fmt, self.sheet[0:2,0:2].number_format)
        self.assertEqual(fmt, self.sheet[0,0].number_format)

    def test_time_format(self):
        fmt = self.document.get_locale().format(pyoo.FORMAT_TIME)
        self.sheet[0:2,0:2].number_format = fmt
        self.assertEqual(fmt, self.sheet[0:2,0:2].number_format)
        self.assertEqual(fmt, self.sheet[0,0].number_format)

    def test_datetime_format(self):
        fmt = self.document.get_locale().format(pyoo.FORMAT_DATETIME)
        self.sheet[0:2,0:2].number_format = fmt
        self.assertEqual(fmt, self.sheet[0:2,0:2].number_format)
        self.assertEqual(fmt, self.sheet[0,0].number_format)


class ChartsTestCase(BaseTestCase):

    _chart_index = 0

    def setUp(self):
        self.sheet = self.document.sheets[0]

    @contextlib.contextmanager
    def create_chart(self, name=None, position=None, ranges=None, **kwargs):
        self.__class__._chart_index += 1
        name = name or 'Chart %d' % self.__class__._chart_index
        position = position or pyoo.SheetPosition(0, 0, 1000, 1000)
        ranges = ranges or pyoo.SheetAddress(0, 0, 2, 3)
        yield self.sheet.charts.create(name, position, ranges, **kwargs)
        del self.sheet.charts[name]

    def test_get_sheet_by_negative_index(self):
        with self.assertRaises(IndexError):
            self.sheet.charts[-1]

    def test_get_sheet_by_too_large_index(self):
        length = len(self.sheet.charts)
        with self.assertRaises(IndexError):
            self.sheet.charts[length]

    def test_get_sheet_by_missing_name(self):
        with self.assertRaises(KeyError):
            self.sheet.charts['Missing']

    def test_create_chart(self):
        with self.create_chart(name='Created chart') as chart:
            self.assertEqual('Created chart', chart.name)
            self.assertEqual(['$A$1:$C$2'], map(str, chart.ranges))

    def test_create_chart_w_multiple_ranges(self):
        ranges = [pyoo.SheetAddress(0, 0, 2, 3), pyoo.SheetAddress(10, 10, 2, 3)]
        with self.create_chart(ranges=ranges) as chart:
            self.assertEqual(['$A$1:$C$2', '$K$11:$M$12'], map(str, chart.ranges))

    def test_create_chart_w_row_header(self):
        with self.create_chart(row_header=True) as chart:
            self.assertTrue(chart.has_row_header)
            self.assertFalse(chart.has_col_header)

    def test_create_chart_w_col_header(self):
        with self.create_chart(col_header=True) as chart:
            self.assertFalse(chart.has_row_header)
            self.assertTrue(chart.has_col_header)

    def test_create_chart_w_cells_as_position(self):
        position = self.sheet[:10,:5]
        with self.create_chart(name='Chart over cells', position=position) as chart:
            self.assertEqual('Chart over cells', chart.name)

    def test_create_chart_w_cells_as_address(self):
        address = self.sheet[:2,:3]
        with self.create_chart(ranges=address) as chart:
            self.assertEqual(['$A$1:$C$2'], map(str, chart.ranges))

    def test_default_diagram_type(self):
        with self.create_chart() as chart:
            self.assertIsInstance(chart.diagram, pyoo.BarDiagram)

    def test_change_diagram_type(self):
        with self.create_chart() as chart:
            chart.change_type(pyoo.LineDiagram)
            self.assertIsInstance(chart.diagram, pyoo.LineDiagram)

    def test_stacked_diagram(self):
        with self.create_chart() as chart:
            diagram = chart.diagram
            self.assertFalse(diagram.is_stacked)
            diagram.is_stacked = True
            self.assertTrue(diagram.is_stacked)

    def test_diagram_spline(self):
        with self.create_chart() as chart:
            diagram = chart.change_type(pyoo.LineDiagram)
            self.assertFalse(diagram.spline)
            diagram.spline = True
            self.assertTrue(diagram.spline)

    def test_secondary_x_axis(self):
        with self.create_chart() as chart:
            diagram = chart.diagram
            self.assertFalse(diagram.has_secondary_x_axis)
            diagram.has_secondary_x_axis = True
            self.assertTrue(diagram.has_secondary_x_axis)

    def test_secondary_y_axis(self):
        with self.create_chart() as chart:
            diagram = chart.diagram
            self.assertFalse(diagram.has_secondary_y_axis)
            diagram.has_secondary_y_axis = True
            self.assertTrue(diagram.has_secondary_y_axis)

    def test_series_too_large_index(self):
        with self.create_chart() as chart:
            with self.assertRaises(IndexError):
                chart.diagram.series[3]

    def test_series_axis(self):
        with self.create_chart() as chart:
            series = chart.diagram.series[0]
            self.assertEqual(pyoo.AXIS_PRIMARY, series.axis)
            series.axis = pyoo.AXIS_SECONDARY
            self.assertEqual(pyoo.AXIS_SECONDARY, series.axis)

    def test_series_line_color(self):
        with self.create_chart() as chart:
            series = chart.diagram.series[0]
            series.line_color = 0xFF0000
            self.assertEqual(0xFF0000, series.line_color)

    def test_series_fill_color(self):
        with self.create_chart() as chart:
            series = chart.diagram.series[0]
            series.fill_color = 0xFF0000
            self.assertEqual(0xFF0000, series.fill_color)


class SpreadsheetCollectionTestCase(BaseTestCase):

    def test_get_sheet_by_invalid_type(self):
        with self.assertRaises(TypeError):
            self.document.sheets[object()]

    def test_get_sheet(self):
        index = self.document.sheets['Sheet1'].name
        self.assertEqual('Sheet1', self.document.sheets[index].name)

    def test_get_sheet_by_negative_index(self):
        with self.assertRaises(IndexError):
            self.document.sheets[-1]

    def test_get_sheet_by_too_large_index(self):
        length = len(self.document.sheets)
        with self.assertRaises(IndexError):
            self.document.sheets[length]

    def test_get_sheet_by_missing_name(self):
        with self.assertRaises(KeyError):
            self.document.sheets['Missing']

    def test_create_sheet(self):
        sheet = self.document.sheets.create('Created')
        self.assertEqual('Created', sheet.name)
        self.assertEqual(len(self.document.sheets) - 1, sheet.index)

    def test_create_sheet_with_index(self):
        sheet = self.document.sheets.create('Created with index', 0)
        self.assertEqual('Created with index', sheet.name)
        self.assertEqual(0, sheet.index)

    def test_copy_sheet(self):
        sheet = self.document.sheets.copy(self.document.sheets[0].name, 'Copied')
        self.assertEqual('Copied', sheet.name)
        self.assertEqual(len(self.document.sheets) - 1, sheet.index)

    def test_copy_sheet_with_index(self):
        sheet = self.document.sheets.copy(self.document.sheets[0].name, 'Copied with index', 0)
        self.assertEqual('Copied with index', sheet.name)
        self.assertEqual(0, sheet.index)

    def test_del_sheet_by_index(self):
        length = len(self.document.sheets)
        sheet = self.document.sheets.create('To delete by name')
        del self.document.sheets[sheet.index]
        self.assertEqual(length, len(self.document.sheets))

    def test_del_sheet_by_name(self):
        length = len(self.document.sheets)
        sheet = self.document.sheets.create('To delete')
        del self.document.sheets[sheet.name]
        self.assertEqual(length, len(self.document.sheets))

    def test_del_sheet_by_missing_index(self):
        length = len(self.document.sheets)
        with self.assertRaises(IndexError):
            del self.document.sheets[length]

    def test_del_sheet_by_missing_name(self):
        with self.assertRaises(KeyError):
            del self.document.sheets['Missing']

    def test_sheet_unicode(self):
        sheet = self.document.sheets[0]
        sheet.name = 'My Sheet'
        self.assertEqual(u'My Sheet', unicode(sheet))

    def test_sheet_str(self):
        sheet = self.document.sheets[0]
        sheet.name = 'My Sheet'
        self.assertEqual('My Sheet', str(sheet))

    def test_sheet_repr(self):
        sheet = self.document.sheets[0]
        sheet.name = 'My Sheet'
        self.assertEqual("<Sheet: 'My Sheet'>", repr(sheet))

    def test_sheet_index(self):
        self.assertEqual(0, self.document.sheets[0].index)


if __name__ == '__main__':
    unittest2.main()
