# -*- coding: UTF-8 -*-
"""
PyOO - Pythonic interface to Apache OpenOffice API (UNO)

Copyright (c) 2016 Seznam.cz, a.s.

"""

import contextlib
import datetime
import unittest

import pyoo


desktop = None


def setUpModule():
    global desktop
    desktop = pyoo.Desktop()

@pyoo.str_repr
class MyObject(object):
    def __str__(self):
        return u'my object'


class SheetPositionTestCase(unittest.TestCase):

    if pyoo.PY2:
        def test_point_str(self):
            position = pyoo.SheetPosition(10, 20)
            self.assertEqual(b'x=10, y=20', str(position))
            self.assertIsInstance(str(position), str)

    def test_point_text(self):
        position = pyoo.SheetPosition(10, 20)
        self.assertEqual(u'x=10, y=20', pyoo.text_type(position))
        self.assertIsInstance(pyoo.text_type(position), pyoo.text_type)

    if pyoo.PY2:
        def test_rectange_str(self):
            position = pyoo.SheetPosition(10, 20, 30, 40)
            self.assertEqual(b'x=10, y=20, width=30, height=40', str(position))
            self.assertIsInstance(str(position), str)

    def test_rectange_text(self):
        position = pyoo.SheetPosition(10, 20, 30, 40)
        self.assertEqual(u'x=10, y=20, width=30, height=40', pyoo.text_type(position))
        self.assertIsInstance(pyoo.text_type(position), pyoo.text_type)

    def test_replace_x(self):
        position = pyoo.SheetPosition(10, 20, 30, 40)
        self.assertEqual('x=100, y=20, width=30, height=40', str(position.replace(x=100)))

    def test_replace_y(self):
        position = pyoo.SheetPosition(10, 20, 30, 40)
        self.assertEqual('x=10, y=200, width=30, height=40', str(position.replace(y=200)))

    def test_replace_width(self):
        position = pyoo.SheetPosition(10, 20, 30, 40)
        self.assertEqual('x=10, y=20, width=300, height=40', str(position.replace(width=300)))

    def test_replace_height(self):
        position = pyoo.SheetPosition(10, 20, 30, 40)
        self.assertEqual('x=10, y=20, width=30, height=400', str(position.replace(height=400)))


class SheetAddressTestCase(unittest.TestCase):

    if pyoo.PY2:
        def test_cell_str(self):
            address = pyoo.SheetAddress(0, 1)
            self.assertEqual(b'$B$1', str(address))
            self.assertIsInstance(str(address), str)

    def test_cell_text(self):
        address = pyoo.SheetAddress(0, 1)
        self.assertEqual(u'$B$1', pyoo.text_type(address))
        self.assertIsInstance(pyoo.text_type(address), pyoo.text_type)

    if pyoo.PY2:
        def test_ranges_str(self):
            address = pyoo.SheetAddress(0, 1, 2, 3)
            self.assertEqual(b'$B$1:$D$2', str(address))
            self.assertIsInstance(str(address), str)

    def test_range_text(self):
        address = pyoo.SheetAddress(0, 1, 2, 3)
        self.assertEqual('$B$1:$D$2', pyoo.text_type(address))
        self.assertIsInstance(pyoo.text_type(address), pyoo.text_type)

    def test_cell_formula(self):
        address = pyoo.SheetAddress(0, 1)
        self.assertEqual('B1', address.formula())

    def test_cell_formula_row_abs(self):
        address = pyoo.SheetAddress(0, 1)
        self.assertEqual('B$1', address.formula(row_abs=True))

    def test_cell_formula_col_abs(self):
        address = pyoo.SheetAddress(0, 1)
        self.assertEqual('$B1', address.formula(col_abs=True))

    def test_cell_formula_row_and_col_abs(self):
        address = pyoo.SheetAddress(0, 1)
        self.assertEqual('$B$1', address.formula(row_abs=True, col_abs=True))

    def test_range_formula(self):
        address = pyoo.SheetAddress(0, 1, 2, 3)
        self.assertEqual('B1:D2', address.formula())

    def test_range_formula_row_abs(self):
        address = pyoo.SheetAddress(0, 1, 2, 3)
        self.assertEqual('B$1:D$2', address.formula(row_abs=True))

    def test_range_formula_col_abs(self):
        address = pyoo.SheetAddress(0, 1, 2, 3)
        self.assertEqual('$B1:$D2', address.formula(col_abs=True))

    def test_range_formula_row_and_col_abs(self):
        address = pyoo.SheetAddress(0, 1, 2, 3)
        self.assertEqual('$B$1:$D$2', address.formula(row_abs=True, col_abs=True))

    def test_replace_row(self):
        address = pyoo.SheetAddress(0, 1, 2, 3)
        self.assertEqual('$B$2:$D$3', str(address.replace(row=1)))

    def test_replace_col(self):
        address = pyoo.SheetAddress(0, 1, 2, 3)
        self.assertEqual('$C$1:$E$2', str(address.replace(col=2)))

    def test_replace_row_count(self):
        address = pyoo.SheetAddress(0, 1, 2, 3)
        self.assertEqual('$B$1:$D$1', str(address.replace(row_count=1)))

    def test_replace_col_count(self):
        address = pyoo.SheetAddress(0, 1, 2, 3)
        self.assertEqual('$B$1:$B$2', str(address.replace(col_count=1)))



class BaseDocumentTestCase(unittest.TestCase):
    """
    Base class for test cases which require spreadsheet document.
    """

    @classmethod
    def setUpClass(cls):
        cls.document = desktop.create_spreadsheet()

    @classmethod
    def tearDownClass(cls):
        cls.document.close()


class CellRangeTestCase(BaseDocumentTestCase):

    def setUp(self):
        self.sheet = self.document.sheets[0]

    # Test conversion to string

    def test_cell_text(self):
        cell = self.sheet[0,0]
        self.assertEqual(u'$A$1', pyoo.text_type(cell))
        self.assertEqual(u'$A$1', pyoo.text_type(cell.address))

    if pyoo.PY2:
        def test_cell_str(self):
            cell = self.sheet[0,0]
            self.assertEqual(b'$A$1', str(cell))
            self.assertEqual(b'$A$1', str(cell.address))

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

    def test_empty_cell_value(self):
        cell = self.sheet[0, 0]
        cell.value = None
        self.assertEqual(None, cell.value)
        self.assertEqual('#N/A', cell.formula)
        self.assertEqual(None, cell.date)
        self.assertEqual(None, cell.time)

    def test_int_cell_value(self):
        cell = self.sheet[0, 0]
        cell.value = 1
        self.assertEqual(1, cell.value)
        self.assertEqual('1', cell.formula)

    def test_positive_long_cell_value(self):
        cell = self.sheet[0, 0]
        cell.value = 2147483648
        self.assertEqual(2147483648, cell.value)
        self.assertEqual('2147483648', cell.formula)

    def test_negative_long_cell_value(self):
        cell = self.sheet[0, 0]
        cell.value = -2147483649
        self.assertEqual(-2147483649, cell.value)
        self.assertEqual('-2147483649', cell.formula)

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
        cell.value = datetime.time(23, 55, 1)
        # Should be almost equal because dates and times are represented as floats
        self.assertEqual(datetime.time(23, 55, 1), cell.time.replace(microsecond=0))

    def test_object_cell_value(self):
        cell = self.sheet[0, 0]
        cell.value = MyObject()
        self.assertEqual(u'my object', cell.value)

    def test_empty_cell_formula(self):
        cell = self.sheet[0, 0]
        cell.formula = None
        self.assertEqual('', cell.value)
        self.assertEqual('', cell.formula)

    def test_int_cell_formula(self):
        cell = self.sheet[0, 0]
        cell.formula = 1
        self.assertEqual(1, cell.value)
        self.assertEqual('1', cell.formula)

    def test_positive_long_cell_formula(self):
        cell = self.sheet[0, 0]
        cell.formula = 2147483648
        self.assertEqual(2147483648, cell.value)
        self.assertEqual('2147483648', cell.formula)

    def test_negative_long_cell_formula(self):
        cell = self.sheet[0, 0]
        cell.formula = -2147483649
        self.assertEqual(-2147483649, cell.value)
        self.assertEqual('-2147483649', cell.formula)

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

    def test_object_cell_formula(self):
        cell = self.sheet[0, 0]
        cell.formula = MyObject()
        self.assertEqual(u'my object', cell.value)
        self.assertEqual(u'my object', cell.formula)

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
        cell.formula = datetime.time(23, 55, 1)
        # Should be almost equal because dates and times are represented as floats
        self.assertEqual(datetime.time(23, 55, 1), cell.time.replace(microsecond=0))

    # Test data access:

    def test_tabular_range_data(self):
        data = [[i * j for j in range(5)] for i in range(10)]
        self.sheet[10:20,1:6].values = data
        self.assertEqual(36, self.sheet[19,5].value)
        self.assertEqual(36, self.sheet[10:20,1:6].values[-1][-1])

    def test_horizontal_range_data(self):
        data = [i + 100 for i in range(5)]
        self.sheet[10,1:6].values = data
        self.assertEqual(104, self.sheet[10,5].value)
        self.assertEqual(104, self.sheet[10,1:6].values[-1])

    def test_vertical_range_data(self):
        data = [i + 200 for i in range(10)]
        self.sheet[10:20,1].values = data
        self.assertEqual(209, self.sheet[19,1].value)
        self.assertEqual(209, self.sheet[10:20,1].values[-1])

    def test_cell_data(self):
        self.sheet[10,1].value = 300
        self.assertEqual(300, self.sheet[10,1].value)

    def test_tabular_range_formulas(self):
        formulas = [['=%d*%d' % (i, j) for j in range(5)] for i in range(10)]
        self.sheet[10:20,1:6].formulas = formulas
        self.assertEqual('=9*4', self.sheet[19,5].formula)
        self.assertEqual('=9*4', self.sheet[10:20,1:6].formulas[-1][-1])

    def test_horizontal_range_formulas(self):
        formulas = ['=%d+100' % i for i in range(5)]
        self.sheet[10,1:6].formulas = formulas
        self.assertEqual('=4+100', self.sheet[10,5].formula)
        self.assertEqual('=4+100', self.sheet[10,1:6].formulas[-1])

    def test_vertical_range_formulas(self):
        formulas = ['=%d+200' % i for i in range(10)]
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

    def test_border_left_width(self):
        cells = self.sheet[20:30,20:30]
        cells.border_left_width = 100
        # Border widths are approximate
        self.assertAlmostEqual(100, cells.border_left_width, delta=10)

    def test_border_right_width(self):
        cells = self.sheet[20:30,20:30]
        cells.border_right_width = 100
        # Border widths are approximate
        self.assertAlmostEqual(100, cells.border_right_width, delta=10)

    def test_border_top_width(self):
        cells = self.sheet[20:30,20:30]
        cells.border_top_width = 100
        # Border widths are approximate
        self.assertAlmostEqual(100, cells.border_top_width, delta=10)

    def test_border_bottom_width(self):
        cells = self.sheet[20:30,20:30]
        cells.border_bottom_width = 100
        # Border widths are approximate
        self.assertAlmostEqual(100, cells.border_bottom_width, delta=10)


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


class ChartsTestCase(BaseDocumentTestCase):

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
            self.assertEqual(['$A$1:$C$2'], list(map(str, chart.ranges)))

    def test_create_chart_w_multiple_ranges(self):
        ranges = [pyoo.SheetAddress(0, 0, 2, 3), pyoo.SheetAddress(10, 10, 2, 3)]
        with self.create_chart(ranges=ranges) as chart:
            self.assertEqual(['$A$1:$C$2', '$K$11:$M$12'], list(map(str, chart.ranges)))

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
            self.assertEqual(['$A$1:$C$2'], list(map(str, chart.ranges)))

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

    def test_vertical_bar_chart(self):
        with self.create_chart() as chart:
            diagram = chart.diagram
            self.assertFalse(diagram.is_horizontal)
            diagram.is_horizontal = True
            self.assertTrue(diagram.is_horizontal)

    def test_bar_chart_with_lines(self):
        with self.create_chart() as chart:
            diagram = chart.diagram
            self.assertEqual(0, diagram.lines)
            diagram.lines = 1
            self.assertEqual(1, diagram.lines)

    def test_grouped_bar_chart(self):
        with self.create_chart() as chart:
            diagram = chart.diagram
            diagram.is_grouped = False
            self.assertFalse(diagram.is_grouped)
            diagram.is_grouped = True
            self.assertTrue(diagram.is_grouped)

    def test_diagram_spline(self):
        with self.create_chart() as chart:
            diagram = chart.change_type(pyoo.LineDiagram)
            self.assertFalse(diagram.spline)
            diagram.spline = True
            self.assertTrue(diagram.spline)

    def test_x_axis(self):
        with self.create_chart() as chart:
            diagram = chart.diagram
            self.assertTrue(diagram.x_axis.visible)
            diagram.x_axis.visible = False
            self.assertFalse(diagram.x_axis.visible)

    def test_y_axis(self):
        with self.create_chart() as chart:
            diagram = chart.diagram
            self.assertTrue(diagram.y_axis.visible)
            diagram.y_axis.visible = False
            self.assertFalse(diagram.y_axis.visible)

    def test_secondary_x_axis(self):
        with self.create_chart() as chart:
            diagram = chart.diagram
            self.assertFalse(diagram.secondary_x_axis.visible)
            diagram.secondary_x_axis.visible = True
            self.assertTrue(diagram.secondary_x_axis.visible)

    def test_secondary_y_axis(self):
        with self.create_chart() as chart:
            diagram = chart.diagram
            self.assertFalse(diagram.secondary_y_axis.visible)
            diagram.secondary_y_axis.visible = True
            self.assertTrue(diagram.secondary_y_axis.visible)

    def test_x_axis_title(self):
        with self.create_chart() as chart:
            diagram = chart.diagram
            self.assertEqual('', diagram.x_axis.title)
            diagram.x_axis.title = 'My Title'
            self.assertEqual('My Title', diagram.x_axis.title)

    def test_x_axis_title_given_as_object(self):
        with self.create_chart() as chart:
            diagram = chart.diagram
            self.assertEqual('', diagram.x_axis.title)
            diagram.x_axis.title = MyObject()
            self.assertEqual(u'my object', diagram.x_axis.title)

    def test_y_axis_title(self):
        with self.create_chart() as chart:
            diagram = chart.diagram
            self.assertEqual('', diagram.y_axis.title)
            diagram.y_axis.title = 'My Title'
            self.assertEqual('My Title', diagram.y_axis.title)

    def test_y_axis_title_given_as_object(self):
        with self.create_chart() as chart:
            diagram = chart.diagram
            self.assertEqual('', diagram.y_axis.title)
            diagram.y_axis.title = MyObject()
            self.assertEqual('my object', diagram.y_axis.title)

    def test_secondary_x_axis_title(self):
        with self.create_chart() as chart:
            diagram = chart.diagram
            self.assertEqual('', diagram.secondary_x_axis.title)
            diagram.secondary_x_axis.title = 'My Title'
            self.assertEqual('My Title', diagram.secondary_x_axis.title)

    def test_secondary_y_axis_title(self):
        with self.create_chart() as chart:
            diagram = chart.diagram
            self.assertEqual('', diagram.secondary_y_axis.title)
            diagram.secondary_y_axis.title = 'My Title'
            self.assertEqual('My Title', diagram.secondary_y_axis.title)

    def test_reversed_x_axis(self):
        with self.create_chart() as chart:
            diagram = chart.diagram
            self.assertFalse(diagram.x_axis.reversed)
            diagram.x_axis.reversed = True
            self.assertTrue(diagram.x_axis.reversed)

    def test_reversed_y_axis(self):
        with self.create_chart() as chart:
            diagram = chart.diagram
            self.assertFalse(diagram.y_axis.reversed)
            diagram.y_axis.reversed = True
            self.assertTrue(diagram.y_axis.reversed)

    def test_reversed_secondary_x_axis(self):
        with self.create_chart() as chart:
            diagram = chart.diagram
            self.assertFalse(diagram.secondary_x_axis.reversed)
            diagram.secondary_x_axis.reversed = True
            self.assertTrue(diagram.secondary_x_axis.reversed)

    def test_reversed_secondary_y_axis(self):
        with self.create_chart() as chart:
            diagram = chart.diagram
            self.assertFalse(diagram.secondary_y_axis.reversed)
            diagram.secondary_y_axis.reversed = True
            self.assertTrue(diagram.secondary_y_axis.reversed)

    def test_logarithmic_x_axis(self):
        with self.create_chart() as chart:
            diagram = chart.diagram
            self.assertFalse(diagram.x_axis.logarithmic)
            diagram.x_axis.logarithmic = True
            self.assertTrue(diagram.x_axis.logarithmic)

    def test_logarithmic_y_axis(self):
        with self.create_chart() as chart:
            diagram = chart.diagram
            self.assertFalse(diagram.y_axis.logarithmic)
            diagram.y_axis.logarithmic = True
            self.assertTrue(diagram.y_axis.logarithmic)

    def test_logarithmic_secondary_x_axis(self):
        with self.create_chart() as chart:
            diagram = chart.diagram
            self.assertFalse(diagram.secondary_x_axis.logarithmic)
            diagram.secondary_x_axis.logarithmic = True
            self.assertTrue(diagram.secondary_x_axis.logarithmic)

    def test_logarithmic_secondary_y_axis(self):
        with self.create_chart() as chart:
            diagram = chart.diagram
            self.assertFalse(diagram.secondary_y_axis.logarithmic)
            diagram.secondary_y_axis.logarithmic = True
            self.assertTrue(diagram.secondary_y_axis.logarithmic)

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
            self.assertEqual(0x000000, series.line_color)
            series.line_color = 0xFF0000
            self.assertEqual(0xFF0000, series.line_color,
                "Setting line color of diagram series is sometimes ignored"
                " so this test sometimes fails.")

    def test_series_fill_color(self):
        with self.create_chart() as chart:
            series = chart.diagram.series[0]
            series.fill_color = 0xFF0000
            self.assertEqual(0xFF0000, series.fill_color)


class SpreadsheetCollectionTestCase(BaseDocumentTestCase):

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

    def test_sheet_text(self):
        sheet = self.document.sheets[0]
        sheet.name = 'My Sheet'
        self.assertEqual(u'My Sheet', pyoo.text_type(sheet))

    if pyoo.PY2:
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


class NameGeneratorTestCase(unittest.TestCase):

    def test_empty_name(self):
        get_name = pyoo.NameGenerator()
        self.assertEqual(get_name(''), '1')

    def test_multiple_empty_names(self):
        get_name = pyoo.NameGenerator()
        get_name('')
        self.assertEqual(get_name(''), '2')

    def test_valid_name(self):
        get_name = pyoo.NameGenerator()
        self.assertEqual(get_name('hello'), 'hello')

    def test_suffix_is_added_for_non_unique_names(self):
        get_name = pyoo.NameGenerator()
        get_name('hello')
        self.assertEqual(get_name('hello'), 'hello 2')

    def test_names_are_unique_after_suffix_is_added(self):
        get_name = pyoo.NameGenerator()
        get_name('hello')
        get_name('hello 2')
        self.assertEqual(get_name('hello'), 'hello 3')

    def test_invalid_chars_are_replaced(self):
        get_name = pyoo.NameGenerator()
        self.assertEqual(get_name('hello[]*?:\/'), 'hello')

    def test_name_with_invalid_chars_only(self):
        get_name = pyoo.NameGenerator()
        self.assertEqual(get_name('[]*?:\/'), '1')

    def test_multiple_names_with_invalid_chars_only(self):
        get_name = pyoo.NameGenerator()
        get_name('[]*?:\/')
        self.assertEqual(get_name('[]*?:\/'), '2')

    def test_names_are_trimmed_to_31_chars(self):
        get_name = pyoo.NameGenerator()
        long_name = '1234567890123456789012345678901234567890'
        self.assertEqual(get_name(long_name), '1234567890123456789012345678901')

    def test_names_are_unique_after_trimmed_to_31(self):
        get_name = pyoo.NameGenerator()
        long_name = '1234567890123456789012345678901234567890'
        get_name(long_name)
        self.assertEqual(get_name(long_name), '12345678901234567890123456789 2')

    def test_names_are_trimmed_to_31_even_if_counter_has_two_digits(self):
        get_name = pyoo.NameGenerator()
        long_name = '1234567890123456789012345678901234567890'
        for i in range(9):
            get_name(long_name)
        self.assertEqual(get_name(long_name), '1234567890123456789012345678 10')

    def test_names_with_casesensitive_chars_are_unique(self):
        get_name = pyoo.NameGenerator()
        get_name(u'Test Č')
        self.assertEqual(u'test č 2', get_name(u'test č'))

if __name__ == '__main__':
    unittest.main()
