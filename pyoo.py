
import os
import itertools
import datetime
import numbers


import uno

# Filters used when saving document.
FILTER_PDF_EXPORT = 'writer_pdf_Export'
FILTER_EXCEL_97 = 'MS Excel 97'
FILTER_EXCEL_2007 = 'Calc MS Excel 2007 XML'

# Number format choices
FORMAT_TEXT = uno.getConstantByName('com.sun.star.i18n.NumberFormatIndex.TEXT')
FORMAT_INT = uno.getConstantByName('com.sun.star.i18n.NumberFormatIndex.NUMBER_INT')
FORMAT_FLOAT = uno.getConstantByName('com.sun.star.i18n.NumberFormatIndex.NUMBER_DEC2')
FORMAT_INT_SEP = uno.getConstantByName('com.sun.star.i18n.NumberFormatIndex.NUMBER_1000INT')
FORMAT_FLOAT_SEP = uno.getConstantByName('com.sun.star.i18n.NumberFormatIndex.NUMBER_1000DEC2')
FORMAT_PERCENT_INT = uno.getConstantByName('com.sun.star.i18n.NumberFormatIndex.PERCENT_INT')
FORMAT_PERCENT_FLOAT = uno.getConstantByName('com.sun.star.i18n.NumberFormatIndex.PERCENT_DEC2')
FORMAT_DATE = uno.getConstantByName('com.sun.star.i18n.NumberFormatIndex.DATE_SYSTEM_SHORT')
FORMAT_TIME = uno.getConstantByName('com.sun.star.i18n.NumberFormatIndex.TIME_HHMM')
FORMAT_DATETIME = uno.getConstantByName('com.sun.star.i18n.NumberFormatIndex.DATETIME_SYSTEM_SHORT_HHMM')

# Font weight choices
FONT_WEIGHT_DONTKNOW = uno.getConstantByName('com.sun.star.awt.FontWeight.DONTKNOW')
FONT_WEIGHT_THIN = uno.getConstantByName('com.sun.star.awt.FontWeight.THIN')
FONT_WEIGHT_ULTRALIGHT = uno.getConstantByName('com.sun.star.awt.FontWeight.ULTRALIGHT')
FONT_WEIGHT_LIGHT = uno.getConstantByName('com.sun.star.awt.FontWeight.LIGHT')
FONT_WEIGHT_SEMILIGHT = uno.getConstantByName('com.sun.star.awt.FontWeight.SEMILIGHT')
FONT_WEIGHT_NORMAL = uno.getConstantByName('com.sun.star.awt.FontWeight.NORMAL')
FONT_WEIGHT_SEMIBOLD = uno.getConstantByName('com.sun.star.awt.FontWeight.SEMIBOLD')
FONT_WEIGHT_BOLD = uno.getConstantByName('com.sun.star.awt.FontWeight.BOLD')
FONT_WEIGHT_ULTRABOLD = uno.getConstantByName('com.sun.star.awt.FontWeight.ULTRABOLD')
FONT_WEIGHT_BLACK = uno.getConstantByName('com.sun.star.awt.FontWeight.BLACK')

# Text underline choices (only first three are present here)
UNDERLINE_NONE = uno.getConstantByName('com.sun.star.awt.FontUnderline.NONE')
UNDERLINE_SINGLE = uno.getConstantByName('com.sun.star.awt.FontUnderline.SINGLE')
UNDERLINE_DOUBLE = uno.getConstantByName('com.sun.star.awt.FontUnderline.DOUBLE')

# Text alignment choices
TEXT_ALIGN_STANDARD = 'STANDARD'
TEXT_ALIGN_LEFT = 'LEFT'
TEXT_ALIGN_CENTER = 'CENTER'
TEXT_ALIGN_RIGHT = 'RIGHT'
TEXT_ALIGN_BLOCK = 'BLOCK'
TEXT_ALIGN_REPEAT = 'REPEAT'


# Exceptions thrown by UNO.
# We try to catch them and re-throw Python standard exceptions.
_IndexOutOfBoundsException = uno.getClass('com.sun.star.lang.IndexOutOfBoundsException')
_NoSuchElementException = uno.getClass('com.sun.star.container.NoSuchElementException')
_IOException = uno.getClass('com.sun.star.io.IOException')

_NoConnectException = uno.getClass('com.sun.star.connection.NoConnectException')
_ConnectionSetupException = uno.getClass('com.sun.star.connection.ConnectionSetupException')



class ConnectionError(Exception):
    """
    Unable to connect to UNO API.
    """


def _clean_slice(key, length):
    """
    Validates and normalizes cell range slice.

    >>> _clean_slice(slice(None, None), 10)
    (0, 10)
    >>> _clean_slice(slice(-10, 10), 10)
    (0, 10)
    >>> _clean_slice(slice(-11, 11), 10)
    (0, 10)
    >>> _clean_slice(slice('x', 'y'), 10)
    Traceback (most recent call last):
    ...
    TypeError: Cell indices must be integers, str given.
    >>> _clean_slice(slice(0, 10, 2), 10)
    Traceback (most recent call last):
    ...
    NotImplementedError: Cell slice with step is not supported.
    """
    if key.step is not None:
        raise NotImplementedError('Cell slice with step is not supported.')
    start, stop = key.start, key.stop
    if start is None:
        start = 0
    if stop is None:
        stop = length
    if not isinstance(start, (int, long)):
        raise TypeError('Cell indices must be integers, %s given.' % type(start).__name__)
    if not isinstance(stop, (int, long)):
        raise TypeError('Cell indices must be integers, %s given.' % type(stop).__name__)
    if start < 0:
        start = start + length
    if stop < 0:
        stop = stop + length
    return max(0, start), min(length, stop)


def _clean_index(key, length):
    """
    Validates and normalizes cell range index

    >>> _clean_index(0, 10)
    0
    >>> _clean_index(-10, 10)
    0
    >>> _clean_index(10, 10)
    Traceback (most recent call last):
    ...
    IndexError: Cell index out of range.
    >>> _clean_index(-11, 10)
    Traceback (most recent call last):
    ...
    IndexError: Cell index out of range.
    >>> _clean_index(None, 10)
    Traceback (most recent call last):
    ...
    TypeError: Cell indices must be integers, NoneType given.
    """
    if not isinstance(key, (int, long)):
        raise TypeError('Cell indices must be integers, %s given.' % type(key).__name__)
    if -length <= key < 0:
        return key + length
    elif 0 <= key < length:
        return key
    else:
        raise IndexError('Cell index out of range.')


def _row_name(index):
    """
    Converts row index to row name.

    >>> _row_name(0)
    '1'
    >>> _row_name(10)
    '11'
    """
    return '%d' % (index + 1)


def _col_name(index):
    """
    Converts column index to column name.

    >>> _col_name(0)
    'A'
    >>> _col_name(26)
    'AA'
    """
    for exp in itertools.count(1):
        limit = 26 ** exp
        if index < limit:
            return ''.join(chr(ord('A') + index / (26 ** i) % 26) for i in xrange(exp-1, -1, -1))
        index -= limit


class SheetPosition(object):
    """
    Position of a rectangular are in a spreadsheet.

    This class represent physical position in 100/th mm,
    see SheetAddress class for logical address of cells.

    >>> position = SheetPosition(1000, 2000)
    >>> print position
    x=1000, y=2000
    >>> position = SheetPosition(1000, 2000, 3000, 4000)
    >>> print position
    x=1000, y=2000, width=3000, height=4000
    """

    def __init__(self, x, y, width=0, height=0):
        self.x = x
        self.y = y
        self.width = width
        self.height = height

    def __unicode__(self):
        if self.width == self.height == 0:
            return 'x=%d, y=%d' % (self.x, self.y)
        return 'x=%d, y=%d, width=%d, height=%d' % (self.x, self.y,
                                                    self.width, self.height)

    def __str__(self):
        return unicode(self).encode('utf-8')

    def __repr__(self):
        return '<%s: %r>' % (self.__class__.__name__, str(self))

    @classmethod
    def _from_uno(cls, position, size):
        return cls(position.X, position.Y, size.Width, size.Height)

    def _to_uno(self):
        return uno.createUnoStruct(
            'com.sun.star.awt.Rectangle', X=self.x, Y=self.y,
            Width=self.width, Height=self.height,
        )


class SheetAddress(object):
    """
    Address of a a cell or rectangular range of cells in a spreadsheet.

    This class represent logical address of cells, see SheetPosition
    class for physical location.

    >>> address = SheetAddress(1, 2)
    >>> print address
    $C$2
    >>> address = SheetAddress(1, 2, 3, 4)
    >>> print address
    $C$2:$F$4
    """

    __slots__ = ('row', 'col', 'row_count', 'col_count')

    def __init__(self, row, col, row_count=1, col_count=1):
        self.row, self.col = row, col
        self.row_count, self.col_count = row_count, col_count

    def __unicode__(self):
        start = u'$%s$%s' % (_col_name(self.col), _row_name(self.row))
        if self.row_count == self.col_count == 1:
            return start
        end = u'$%s$%s' % (_col_name(self.col_end), _row_name(self.row_end))
        return u'%s:%s' % (start, end)

    def __str__(self):
        return unicode(self).encode('utf-8')

    def __repr__(self):
        return '<%s: %r>' % (self.__class__.__name__, str(self))

    @property
    def row_end(self):
        return self.row + self.row_count - 1

    @property
    def col_end(self):
        return self.col + self.col_count - 1

    @classmethod
    def _from_uno(cls, target):
        row_count = target.EndRow - target.StartRow + 1
        col_count = target.EndColumn - target.StartColumn + 1
        return cls(target.StartRow, target.StartColumn, row_count, col_count)

    def _to_uno(self, sheet):
        return uno.createUnoStruct(
            'com.sun.star.table.CellRangeAddress', Sheet=sheet,
            StartColumn=self.col, StartRow=self.col,
            EndColumn=self.col_end, EndRow=self.row_end,
        )


class SheetCursor(object):
    """
    Cursor in spreadsheet sheet.

    Most of spreadsheet operations are done using this cursor
    because cursor movement is much faster then cell range selection.
    """

    __slots__ = ('_target', 'row', 'col', 'row_count', 'col_count',
                 'max_row_count', 'max_col_count')

    def __init__(self, target):
        self._target = target # com.sun.star.sheet.XSheetCellCursor
        ra = self._target.getRangeAddress()
        self.row = 0
        self.col = 0
        self.row_count = ra.EndRow + 1
        self.col_count = ra.EndColumn + 1
        # Default cursor contains all the cells.
        self.max_row_count = self.row_count
        self.max_col_count = self.col_count

    def get_target(self, row, col, row_count, col_count):
        """
        Moves cursor to the specified position and returns in.
        """
        # This method is called for almost any operation so it should
        # be maximally optimized.
        #
        # Any comparison here is negligible to UNO call. It means that we do
        # all possible checks which can prevent unnecessary cursor movement.
        #
        # Generally we need to expand or collapse selection to the desired
        # size and move it to the desired position. But both of these actions
        # can fail if there is not enough space. For this reason we must
        # determine which of the actions has to be done first. In some cases
        # we must even move the cursor twice (cursor movement is faster than
        # selection change).
        #
        target = self._target
        # If not we cannot resize selection now then we must move cursor first.
        if self.row + row_count > self.max_row_count or self.col + col_count > self.max_col_count:
            # Move cursor to the desired position if possible.
            row_delta = row - self.row if row + self.row_count <= self.max_row_count else 0
            col_delta = col - self.col if col + self.col_count <= self.max_col_count else 0
            target.gotoOffset(col_delta, row_delta)
            self.row += row_delta
            self.col += col_delta
        # Resize selection
        if (row_count, col_count) != (self.row_count, self.col_count):
            target.collapseToSize(col_count, row_count)
            self.row_count = row_count
            self.col_count = col_count
        # Move cursor to the desired position
        if (row, col) != (self.row, self.col):
            target.gotoOffset(col - self.col, row - self.row)
            self.row = row
            self.col = col
        return target


class CellRange(object):
    """
    Range of cells in one sheet.

    This is an abstract base class implements cell manipulation functionality.
    """

    __slots__ = ('sheet', 'address')

    def __init__(self, sheet, address):
        self.sheet = sheet
        self.address = address

    def __unicode__(self):
        return unicode(self.address)

    def __str__(self):
        return unicode(self).encode('utf-8')

    def __repr__(self):
        return '<%s: %r>' % (self.__class__.__name__, str(self))

    @property
    def position(self):
        """
        Physical position of this cells.
        """
        target = self._get_target()
        position, size = target.getPropertyValues(('Position', 'Size'))
        return SheetPosition._from_uno(position, size)

    def __get_is_merged(self):
        """
        Gets whether cells are merged.
        """
        return self._get_target().getIsMerged()
    def __set_is_merged(self, value):
        """
        Sets whether cells are merged.
        """
        self._get_target().merge(value)
    is_merged = property(__get_is_merged, __set_is_merged)

    def __get_number_format(self):
        """
        Gets format of numbers in this cells.
        """
        return self._get_target().getPropertyValue('NumberFormat')
    def __set_number_format(self, value):
        """
        Sets format of numbers in this cells.
        """
        self._get_target().setPropertyValue('NumberFormat', value)
    number_format = property(__get_number_format, __set_number_format)

    def __get_text_align(self):
        """
        Gets horizontal alignment.

        Returns one of TEXT_ALIGN_* constants.
        """
        return self._get_target().getPropertyValue('HoriJustify').value
    def __set_text_align(self, value):
        """
        Sets horizontal alignment.

        Accepts TEXT_ALIGN_* constants.
        """
        # The HoriJustify property contains is a struct.
        # We need to get it, update value and then set it back.
        target = self._get_target()
        struct = target.getPropertyValue('HoriJustify')
        struct.value = value
        target.setPropertyValue('HoriJustify', struct)
    text_align = property(__get_text_align, __set_text_align)

    def __get_font_size(self):
        """
        Gets font size.
        """
        return self._get_target().getPropertyValue('CharHeight')
    def __set_font_size(self, value):
        """
        Sets font size.
        """
        return self._get_target().setPropertyValue('CharHeight', value)
    font_size = property(__get_font_size, __set_font_size)

    def __get_font_weight(self):
        """
        Gets font weight.
        """
        return self._get_target().getPropertyValue('CharWeight')
    def __set_font_weight(self, value):
        """
        Sets font weight.
        """
        return self._get_target().setPropertyValue('CharWeight', value)
    font_weight = property(__get_font_weight, __set_font_weight)

    def __get_underline(self):
        """
        Gets text underline.

        Returns UNDERLINE_* constants.
        """
        return self._get_target().getPropertyValue('CharUnderline')
    def __set_underline(self, value):
        """
        Sets text weight.

        Accepts UNDERLINE_* constants.
        """
        return self._get_target().setPropertyValue('CharUnderline', value)
    underline = property(__get_underline, __set_underline)

    def __get_text_color(self):
        """
        Gets text color.

        Color is returned as integer in format 0xAARRGGBB.
        Returns None if no the text color is not set.
        """
        value = self._get_target().getPropertyValue('CharColor')
        if value == -1:
            value = None
        return value
    def __set_text_color(self, value):
        """
        Sets text color.

        Color should be given as an integer in format 0xAARRGGBB.
        Unsets the text color if None value is given.
        """
        if value is None:
            value = -1
        return self._get_target().setPropertyValue('CharColor', value)
    text_color = property(__get_text_color, __set_text_color)

    def __get_background_color(self):
        """
        Gets cell background color.

        Color is returned as integer in format 0xAARRGGBB.
        Returns None if the background color is not set.
        """
        value = self._get_target().getPropertyValue('CellBackColor')
        if value == -1:
            value = None
        return value
    def __set_background_color(self, value):
        """
        Sets cell background color.

        Color should be given as an integer in format 0xAARRGGBB.
        Unsets the background color if None value is given.
        """
        if value is None:
            value = -1
        return self._get_target().setPropertyValue('CellBackColor', value)
    background_color = property(__get_background_color, __set_background_color)

    def __get_border_width(self):
        """
        Gets width of all cell borders (in 1/100 mm).

        Returns 0 if cell borders are different.
        """
        target = self._get_target()
        # Get four borders and test if all of them have same width.
        keys = ('TopBorder', 'RightBorder', 'BottomBorder', 'LeftBorder')
        lines = target.getPropertyValues(keys)
        values = [line.OuterLineWidth for line in lines]
        if any(value != values[0] for value in values):
            return 0
        return values[0]
    def __set_border_width(self, value):
        """
        Sets width of all cell borders (in 1/100 mm).
        """
        target = self._get_target()
        line = uno.createUnoStruct('com.sun.star.table.BorderLine2')
        line.OuterLineWidth = value
        # Set all four borders using one call - this can save even a few seconds
        keys = ('TopBorder', 'RightBorder', 'BottomBorder', 'LeftBorder')
        lines = (line, line, line, line)
        target.setPropertyValues(keys, lines)
    border_width = property(__get_border_width, __set_border_width)

    def __get_inner_border_width(self):
        """
        Gets with of inner border between cells (in 1/100 mm).

        Returns 0 if cell borders are different.
        """
        target = self._get_target()
        tb = target.getPropertyValue('TableBorder')
        horizontal = tb.HorizontalLine.OuterLineWidth
        vertical = tb.VerticalLine.OuterLineWidth
        if horizontal != vertical:
            return 0
        return horizontal
    def __set_inner_border_width(self, value):
        """
        Sets with of inner border between cells (in 1/100 mm).
        """
        target = self._get_target()
        # Inner borders are saved in a TableBorder struct.
        line = uno.createUnoStruct('com.sun.star.table.BorderLine2')
        line.OuterLineWidth = value
        tb = target.getPropertyValue('TableBorder')
        tb.HorizontalLine = tb.VerticalLine = line
        target.setPropertyValue('TableBorder', tb)
    inner_border_width = property(__get_inner_border_width,
                                  __set_inner_border_width)

    # Internal methods:

    def _get_target(self):
        """
        Returns cursor which can be used for most of operations.
        """
        address = self.address
        cursor = self.sheet.cursor
        return cursor.get_target(address.row, address.col,
                                 address.row_count, address.col_count)

    def _clean_value(self, value):
        """
        Validates and converts value before assigning it to a cell.
        """
        if value is None:
            return value
        if isinstance(value, numbers.Real):
            return value
        if isinstance(value, basestring):
            return value
        if isinstance(value, datetime.date):
            return self.sheet.document.date_to_number(value)
        if isinstance(value, datetime.time):
            return self.sheet.document.time_to_number(value)
        raise ValueError(value)

    def _clean_formula(self, value):
        """
        Validates and converts formula before assigning it to a cell.
        """
        if value is None:
            return ''
        if isinstance(value, numbers.Real):
            return value
        if isinstance(value, basestring):
            return value
        if isinstance(value, datetime.date):
            return self.sheet.document.date_to_number(value)
        if isinstance(value, datetime.time):
            return self.sheet.document.time_to_number(value)
        raise ValueError(value)


class Cell(CellRange):
    """
    One cell in a spreadsheet.

    Cells are returned when a sheet (or any other tabular cell range)
    is indexed by two integer numbers.
    """

    __slots__ = ()

    def __get_value(self):
        """
        Gets cell value with as a string or number based on cell type.
        """
        array = self._get_target().getDataArray()
        return array[0][0]
    def __set_value(self, value):
        """
        Sets cell value to a string or number based on the given value.
        """
        array = ((self._clean_value(value),),)
        return self._get_target().setDataArray(array)
    value = property(__get_value, __set_value)

    def __get_formula(self):
        """
        Gets a formula in this cell.

        If this cell contains actual formula then the returned value starts
        with an equal sign but any cell value is returned.
        """
        array = self._get_target().getFormulaArray()
        return array[0][0]
    def __set_formula(self, formula):
        """
        Sets a formula in this cell.

        Any cell value can be set using this method. Actual formulas must
        start with an equal sign.
        """
        array = ((self._clean_formula(formula),),)
        return self._get_target().setFormulaArray(array)
    formula = property(__get_formula, __set_formula)

    @property
    def date(self):
        """
        Returns date value in this cell.

        Converts value from number to datetime.datetime instance.
        """
        return self.sheet.document.date_from_number(self.value)

    @property
    def time(self):
        """
        Returns time value in this cell.

        Converts value from number to datetime.time instance.
        """
        return self.sheet.document.time_from_number(self.value)



class TabularCellRange(CellRange):
    """
    Tabular range of cells.

    Individual cells can be accessed by (row, column) index and
    slice notation can be used for retrieval of sub ranges.

    Instances of this class are returned when a sheet (or any other tabular
    cell range) is sliced in both axes.
    """

    __slots__ = ()

    def __len__(self):
        return self.address.row_count

    def __getitem__(self, key):
        if not isinstance(key, tuple):
            # Expression cells[row] is equal to cells[row, :] and
            # expression cells[start:stop] is equal to cells[start:stop, :].
            key = (key, slice(None))
        elif len(key) != 2:
            raise ValueError('Cell range has two dimensions.')
        address = self.address
        row_val, col_val = key
        if isinstance(row_val, slice):
            start, stop = _clean_slice(row_val, address.row_count)
            row, row_count = address.row + start, stop - start
            single_row = False
        else:
            index = _clean_index(row_val, address.row_count)
            row, row_count = address.row + index, 1
            single_row = True
        if isinstance(col_val, slice):
            start, stop = _clean_slice(col_val, address.col_count)
            col, col_count = address.col + start, stop - start
            single_col = False
        else:
            index = _clean_index(col_val, address.col_count)
            col, col_count = address.col + index, 1
            single_col = True

        address = SheetAddress(row, col, row_count, col_count)
        if single_row and single_col:
            return Cell(self.sheet, address)
        if single_row:
            return HorizontalCellRange(self.sheet, address)
        if single_col:
            return VerticalCellRange(self.sheet, address)
        return TabularCellRange(self.sheet, address)

    def __get_values(self):
        """
        Gets values in this cell range as a tuple of tuples.
        """
        array = self._get_target().getDataArray()
        return array
    def __set_values(self, values):
        """
        Sets values in this cell range from an iterable of iterables.
        """
        # Tuple of tuples is required
        array = tuple(tuple(self._clean_value(col) for col in row) for row in values)
        self._get_target().setDataArray(array)
    values = property(__get_values, __set_values)

    def __get_formulas(self):
        """
        Gets formulas in this cell range as a tuple of tuples.

        If cells contain actual formulas then the returned values start
        with an equal sign  but all values are returned.
        """
        return self._get_target().getFormulaArray()
    def __set_formulas(self, formulas):
        """
        Sets formulas in this cell range from an iterable of iterables.

        Any cell values can be set using this method. Actual formulas must
        start with an equal sign.
        """
        # Tuple of tuples is required
        array = tuple(tuple(self._clean_formula(col) for col in row) for row in formulas)
        self._get_target().setFormulaArray(array)
    formulas = property(__get_formulas, __set_formulas)


class HorizontalCellRange(CellRange):
    """
    Range of cell in one row.

    Individual cells can be accessed by integer index or subranges
    can be retrieved using slice notation.

    Instances of this class are returned if a sheet (or any other tabular
    cell range) is indexed by a row number but columns are sliced.
    """

    __slots__ = ()

    def __len__(self):
        return self.address.col_count

    def __getitem__(self, key):
        if isinstance(key, slice):
            start, stop = _clean_slice(key, len(self))
            address = SheetAddress(self.address.row, self.address.col + start,
                                   self.address.row_count, stop - start)
            return HorizontalCellRange(self.sheet, address)
        else:
            index = _clean_index(key, len(self))
            address = SheetAddress(self.address.row, self.address.col + index)
            return Cell(self.sheet, address)

    def __get_values(self):
        """
        Gets values in this cell range as a tuple.
        """
        array = self._get_target().getDataArray()
        return array[0]
    def __set_values(self, values):
        """
        Sets values in this cell range from an iterable.
        """
        array = (tuple(self._clean_value(v) for v in values),)
        self._get_target().setDataArray(array)
    values = property(__get_values, __set_values)

    def __get_formulas(self):
        """
        Gets formulas in this cell range as a tuple.

        If cells contain actual formulas then the returned values start
        with an equal sign  but all values are returned.
        """
        array = self._get_target().getFormulaArray()
        return array[0]
    def __set_formulas(self, formulas):
        """
        Sets formulas in this cell range from an iterable.

        Any cell values can be set using this method. Actual formulas must
        start with an equal sign.
        """
        array = (tuple(self._clean_formula(v) for v in formulas),)
        return self._get_target().setFormulaArray(array)
    formulas = property(__get_formulas, __set_formulas)


class VerticalCellRange(CellRange):
    """
    Range of cell in one column.

    Individual cells can be accessed by integer index or or subranges
    can be retrieved using slice notation.

    Instances of this class are returned if a sheet (or any other tabular
    cell range) is indexed by a column number but rows are sliced.
    """

    __slots__ = ()

    def __len__(self):
        return self.address.row_count

    def __getitem__(self, key):
        if isinstance(key, slice):
            start, stop = _clean_slice(key, len(self))
            address = SheetAddress(self.address.row  + start, self.address.col,
                                   stop - start, self.address.col_count)
            return HorizontalCellRange(self.sheet, address)
        else:
            index = _clean_index(key, len(self))
            address = SheetAddress(self.address.row  + index, self.address.col)
            return Cell(self.sheet, address)

    def __get_values(self):
        """
        Gets values in this cell range as a tuple.

        This is much more effective than reading cell values one by one.
        """
        array = self._get_target().getDataArray()
        return tuple(itertools.chain.from_iterable(array))
    def __set_values(self, values):
        """
        Sets values in this cell range from an iterable.

        This is much more effective than writing cell values one by one.
        """
        array = tuple((self._clean_value(v),) for v in values)
        self._get_target().setDataArray(array)
    values = property(__get_values, __set_values)

    def __get_formulas(self):
        """
        Gets formulas in this cell range as a tuple.

        If cells contain actual formulas then the returned values start
        with an equal sign  but all values are returned.
        """
        array = self._get_target().getFormulaArray()
        return tuple(itertools.chain.from_iterable(array))
    def __set_formulas(self, formulas):
        """
        Sets formulas in this cell range from an iterable.

        Any cell values can be set using this method. Actual formulas must
        start with an equal sign.
        """
        array = tuple((self._clean_formula(v),) for v in formulas)
        self._get_target().setFormulaArray(array)
    formulas = property(__get_formulas, __set_formulas)


class NamedCollection(object):
    """
    Base class for collections accessible by both index and name.
    """

    __slots__ = ('_target',)

    def __init__(self, target):
        # Target must implement both of:
        # http://www.openoffice.org/api/docs/common/ref/com/sun/star/container/XIndexAccess.html
        # http://www.openoffice.org/api/docs/common/ref/com/sun/star/container/XNameAccess.html
        self._target = target

    def __len__(self):
        return self._target.getCount()

    def __getitem__(self, key):
        if isinstance(key, (int, long)):
            target = self._get_by_index(key)
            return self._factory(target)
        if isinstance(key, basestring):
            target = self._get_by_name(key)
            return self._factory(target)
        raise TypeError('%s must be accessed either by index or name.'
                        % self.__class__.__name__)

    # Internal:

    def _factory(self, target):
        raise NotImplementedError # pragma: no cover

    def _get_by_index(self, index):
        try:
            # http://www.openoffice.org/api/docs/common/ref/com/sun/star/container/XIndexAccess.html#getByIndex
            return self._target.getByIndex(index)
        except _IndexOutOfBoundsException:
            raise IndexError(index)

    def _get_by_name(self, name):
        try:
            # http://www.openoffice.org/api/docs/common/ref/com/sun/star/container/XNameAccess.html#getByName
            return self._target.getByName(name)
        except _NoSuchElementException:
            raise KeyError(name)



class Chart(object):
    """
    Chart
    """

    __slots__ = ('sheet', '_target')

    def __init__(self, sheet, target):
        self.sheet = sheet
        self._target = target

    @property
    def name(self):
        return self._target.getName()

    @property
    def has_row_header(self):
        return self._target.getHasRowHeaders()

    @property
    def has_col_header(self):
        return self._target.getHasColumnHeaders()

    @property
    def ranges(self):
        ranges = self._target.getRanges()
        return map(SheetAddress._from_uno, ranges)


class ChartCollection(NamedCollection):
    """
    Collection of charts in one sheet.
    """

    __slots__ = ('sheet',)

    def __init__(self, sheet, target):
        self.sheet = sheet
        super(ChartCollection, self).__init__(target)

    def create(self, name, position, ranges=(), col_header=False, row_header=False):
        """
        Creates a and inserts a new chart.
        """
        rect = self._uno_rect(position)
        ranges = self._uno_ranges(ranges)
        self._create(name, rect, ranges, col_header, row_header)
        return self[name]

    # Internal:

    def _factory(self, target):
        return Chart(self.sheet, target)

    def _uno_rect(self, position):
        if isinstance(position, CellRange):
            position = position.position
        return position._to_uno()

    def _uno_ranges(self, ranges):
        if not isinstance(ranges, (list, tuple)):
            ranges = [ranges]
        return tuple(map(self._uno_range, ranges))

    def _uno_range(self, address):
        if isinstance(address, CellRange):
            address = address.address
        return address._to_uno(self.sheet.index)

    def _create(self, name, rect, ranges, col_header, row_header):
        # http://www.openoffice.org/api/docs/common/ref/com/sun/star/table/XTableCharts.html#addNewByName
        self._target.addNewByName(name, rect, ranges, col_header, row_header)


class Sheet(TabularCellRange):
    """
    One sheet in a spreadsheet document.

    This class extends TabularCellRange which means that cells can
    be accessed using index or slice notation.

    Sheet instances can be accessed using sheets property
    of a SpreadsheetDocument class.
    """

    __slots__ = ('document', '_target', 'cursor')

    def __init__(self, document, target):
        self.document = document # Parent SpreadsheetDocument.
        self._target = target # UNO com.sun.star.sheet.XSpreadsheet
        # This cursor will be used for most of the operation in this sheet.
        self.cursor = SheetCursor(target.createCursor())
        # Determine size of this sheet using the created cursor.
        address = SheetAddress(0, 0, self.cursor.row_count, self.cursor.col_count)
        super(Sheet, self).__init__(self, address)

    def __unicode__(self):
        return unicode(self.name)

    @property
    def index(self):
        """
        Index of this sheet in the document.
        """
        # This should be cached if used more often.
        return self._target.getRangeAddress().Sheet

    def __get_name(self):
        """
        Gets a name of this sheet.
        """
        # This should be cached if used more often.
        return self._target.getName();
    def __set_name(self, value):
        """
        Sets a name of this sheet.
        """
        return self._target.setName(value);
    name = property(__get_name, __set_name)

    @property
    def charts(self):
        target = self._target.getCharts()
        return ChartCollection(self, target)


class SpreadsheetCollection(NamedCollection):
    """
    Collection of spreadsheets in a spreadsheet document.

    Instance of this class is returned via sheets property of
    the SpreadsheetDocument class.
    """

    def __init__(self, document, target):
        self.document = document # Parent SpreadsheetDocument
        super(SpreadsheetCollection, self).__init__(target)

    def __delitem__(self, key):
        if not isinstance(key, basestring):
            key = self[key].name
        self._delete(key)

    def create(self, name, index=None):
        """
        Creates a new sheet with the given name.

        If an optional index argument is not provided then the created
        sheet is appended at the end. Returns the new sheet.
        """
        if index is None:
            index = len(self)
        self._create(name, index)
        return self[name]

    def copy(self, old_name, new_name, index=None):
        """
        Copies an old sheet with the old_name to a new sheet with new_name.

        If an optional index argument is not provided then the created
        sheet is appended at the end. Returns the new sheet.
        """
        if index is None:
            index = len(self)
        self._copy(old_name, new_name, index)
        return self[new_name]

    # Internal:

    def _factory(self, target):
        return Sheet(self.document, target)

    def _create(self, name, index):
        # http://www.openoffice.org/api/docs/common/ref/com/sun/star/sheet/XSpreadsheets.html#insertNewByName
        self._target.insertNewByName(name, index)

    def _copy(self, old_name, new_name, index):
        # http://www.openoffice.org/api/docs/common/ref/com/sun/star/sheet/XSpreadsheets.html#copyByName
        self._target.copyByName(old_name, new_name, index)

    def _delete(self, name):
        try:
            self._target.removeByName(name)
        except _NoSuchElementException:
            raise KeyError(name)


class Locale(object):
    """
    Document locale.

    Provides locale number formats. Instances of this class can be
    retrieved from SpreadsheetDocument using get_locale method.
    """

    def __init__(self, locale, formats):
        self._locale = locale
        self._formats = formats

    def format(self, code):
        """
        Returns one of predefined formats.

        Accepts FORMAT_* constants.
        """
        # http://www.openoffice.org/api/docs/common/ref/com/sun/star/util/XNumberFormatTypes.html#getFormatIndex
        return self._formats.getFormatIndex(code, self._locale)


class SpreadsheetDocument(object):
    """
    Spreadsheet document.
    """

    def __init__(self, document):
        self._document = document

    def save(self, path, filter_name=None):
        """
        Save document to a local file system.

        Accept optional second  argument which defines type of saved file.
        Use one of FILTER_* constants or see list of available filters at
        http://wakka.net/archives/7
        """
        # UNO requires absolute paths
        url = uno.systemPathToFileUrl(os.path.abspath(path))
        if filter_name:
            format_filter = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
            format_filter.Name = 'FilterName'
            format_filter.Value = filter_name
            filters = (format_filter,)
        else:
            filters = ()
        # http://www.openoffice.org/api/docs/common/ref/com/sun/star/frame/XStorable.html#storeToURL
        try:
            self._document.storeToURL(url, filters)
        except _IOException, e:
            raise IOError(e.Message)

    def close(self):
        """
        Closes this document.
        """
        # http://www.openoffice.org/api/docs/common/ref/com/sun/star/util/XCloseable.html#close
        self._document.close(True)

    def get_locale(self, language=None, country=None, variant=None):
        """
        Returns location which can be used for access to number formats.
        """
        # http://www.openoffice.org/api/docs/common/ref/com/sun/star/lang/Locale.html
        locale = uno.createUnoStruct('com.sun.star.lang.Locale')
        if language:
            locale.Language = language
        if country:
            locale.Country = country
        if variant:
            locale.Variant = variant
        formats = self._document.getNumberFormats()
        return Locale(locale, formats)

    @property
    def sheets(self):
        """
        Collection of sheets in this document.
        """
        # http://www.openoffice.org/api/docs/common/ref/com/sun/star/sheet/XSpreadsheetDocument.html#getSheets
        try:
            return self._sheets
        except AttributeError:
            target = self._document.getSheets()
            self._sheets = SpreadsheetCollection(self, target)
        return self._sheets

    def date_from_number(self, value):
        """
        Converts a float value to corresponding datetime instance.
        """
        if not isinstance(value, numbers.Real):
            return None
        delta = datetime.timedelta(days=value)
        return self._null_date + delta

    def date_to_number(self, date):
        """
        Converts a date or datetime instance to a corresponding float value.
        """
        if isinstance(date, datetime.datetime):
            delta = date - self._null_date
        elif isinstance(date, datetime.date):
            delta = date - self._null_date.date()
        else:
            raise TypeError(date)
        return delta.days + delta.seconds / (24.0 * 60 * 60)

    def time_from_number(self, value):
        """
        Converts a float value to corresponding time instance.
        """
        if not isinstance(value, numbers.Real):
            return None
        delta = datetime.timedelta(days=value)
        minutes, second = divmod(delta.seconds, 60)
        hour, minute = divmod(minutes, 60)
        return datetime.time(hour, minute, second)

    def time_to_number(self, time):
        """
        Converts a time instance to a corresponding float value.
        """
        if not isinstance(time, datetime.time):
            raise TypeError(time)
        return ((time.second / 60.0 + time.minute) / 60.0 + time.hour) / 24.0

    # Internal:

    @property
    def _null_date(self):
        """
        Returns date which is represented by a integer 0.
        """
        # http://www.openoffice.org/api/docs/common/ref/com/sun/star/util/NumberFormatSettings.html#NullDate
        try:
            return self.__null_date
        except AttributeError:
            number_settings = self._document.getNumberFormatSettings()
            d = number_settings.getPropertyValue('NullDate')
            self.__null_date = datetime.datetime(d.Year, d.Month, d.Day)
        return self.__null_date


def _get_connection_url(hostname, port):
    return 'uno:socket,host=%s,port=%d;urp;StarOffice.ComponentContext' % (hostname, port)


class Desktop(object):
    """
    Access to a running to an OpenOffice.org program.

    Allows to create and open of spreadsheet documents.

    Opens a connection to a running OpenOffice.org program when Desktop
    instance is initialized. If the program OpenOffice.org is restarted
    then the connection is lost all subsequent method calls will fail.
    """

    def __init__(self, hostname='localhost', port=2002):
        url = _get_connection_url(hostname, port)
        local_context = uno.getComponentContext()
        resolver = local_context.getServiceManager().createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', local_context)
        try:
            remote_context = resolver.resolve(url)
        except (_NoConnectException, _ConnectionSetupException), e:
            raise ConnectionError(e.Message)
        desktop = remote_context.getServiceManager().createInstanceWithContext("com.sun.star.frame.Desktop", remote_context)
        self._desktop = desktop

    def create_spreadsheet(self):
        """
        Creates a new spreadsheet document.
        """
        url = 'private:factory/scalc'
        document = self._open_url(url)
        return SpreadsheetDocument(document)

    def open_spreadsheet(self, path, as_template=False):
        """
        Opens an exiting spreadsheet document on the local file system.
        """
        extra = ()
        if as_template:
            pv = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
            pv.Name = 'AsTemplate'
            pv.Value = True
            extra += (pv,)
        # UNO requires absolute paths
        url = uno.systemPathToFileUrl(os.path.abspath(path))
        document = self._open_url(url, extra)
        return SpreadsheetDocument(document)

    def _open_url(self, url, extra=()):
        # http://www.openoffice.org/api/docs/common/ref/com/sun/star/frame/XComponentLoader.html#loadComponentFromURL
        try:
            return self._desktop.loadComponentFromURL(url, '_blank', 0, extra)
        except _IOException, e:
            raise IOError(e.Message)


class LazyDesktop(object):
    """
    Lazy access to a running to Open Office program.

    Provides same interface as a Desktop class but creates connection
    to OpenOffice program when necessary. The advantage of this approach
    is that a LazyDesktop instance can recover from a restart of
    the OpenOffice.org program.
    """

    cls = Desktop

    def __init__(self, hostname='localhost', port=2002):
        self.hostname = hostname
        self.port = port

    def create_spreadsheet(self):
        """
        Creates a new spreadsheet document.
        """
        desktop = self.cls(self.hostname, self.port)
        return desktop.create_spreadsheet()

    def open_spreadsheet(self, path, as_template=False):
        """
        Opens an exiting spreadsheet document on the local file system.
        """
        desktop = self.cls(self.hostname, self.port)
        return desktop.open_spreadsheet(path, as_template=as_template)
