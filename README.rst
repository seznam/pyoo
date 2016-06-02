
========================================================
PyOO - Pythonic interface to Apache OpenOffice API (UNO)
========================================================

PyOO allows you to control a running OpenOffice_ or LibreOffice_
program for reading and writing spreadsheet documents.
The library can be used for generating documents in various
formats -- including Microsoft Excel 97 (.xls),
Microsoft Excel 2007 (.xlsx) and PDF.

The main advantage of the PyOO library is that it can use almost any
functionality implemented in OpenOffice / LibreOffice applications.
On the other hand it needs a running process of a office suite
application which is significant overhead.

PyOO uses UNO_ interface via Python-UNO_ bridge. UNO is a
standard interface to a running OpenOffice or LibreOffice
application. Python-UNO provides this interface in Python scripts.
Direct usage of UNO API via Python-UNO can be quite complicated
and even simple tasks require a lot of code. Also many UNO calls
are slow and should be avoided when possible.

PyOO wraps a robust Python-UNO bridge to simple and Pythonic
interface. Under the hood it implements miscellaneous
optimizations which can prevent unnecessary expensive UNO
calls.

Available features:

  * Opening and creation of spreadsheet documents
  * Saving documents to all formats available in OpenOffice
  * Charts and diagrams
  * Sheet access and manipulation
  * Formulas
  * Cell merging
  * Number, text, date, and time values
  * Cell and text formating
  * Number formating
  * Locales

If some important feature missing then the UNO API is always available.


.. _OpenOffice: http://www.openoffice.org/
.. _LibreOffice: http://www.libreoffice.org/
.. _UNO: http://www.openoffice.org/api/docs/common/ref/com/sun/star/module-ix.html
.. _Python-UNO: http://www.openoffice.org/udk/python/python-bridge.html


Requirements
------------

PyOO runs on both Python 2 (2.7+) and Python 3 (3.3+).

The only dependency is the Python-UNO library (imported as a module ``uno``).
It is often installed with the office suite. On Debian based systems it can by
installed as ``python-uno`` or ``python3-uno`` package.

Obviously you will also need OpenOffice or LibreOffice Calc.
On Debian systems it is available as ``libreoffice-calc`` package.


Installation
------------

PyOO library can be installed from PYPI using pip_ (or easy_install)::

    $ pip install pyoo

If you downloaded the code you can install it using the  ``setup.py`` script: ::

    $ python setup.py install

Alternatively you can copy the ``pyoo.py`` file somewhere to your ``PYTHONPATH``.

.. _pip: https://pypi.python.org/pypi/pip


Usage
-----

Starting OpenOffice / LibreOffice
.................................

PyOO requires a running OpenOffice or LibreOffice instance which
it can connect to. On Debian you can start LibreOffice from
a command line using a command similar to: ::

    $ soffice --accept="socket,host=localhost,port=2002;urp;" --norestore --nologo --nodefault # --headless

The LibreOffice will be listening for localhost connection
on port 2002. Alternatively a named pipe can be used: ::

    $ soffice --accept="pipe,name=hello;urp;" --norestore --nologo --nodefault # --headless

If the ``--headless`` option is used then no user interface is
visible even when a document is opened.

For more information run: ::

    $ soffice --help

It is recommended to start directly the ``soffice`` binary.
There can be various scripts (called for example ``libreoffice``)
which will run the ``soffice`` binary but you may not get the
correct PID of the running program.


Accessing documents
...................

PyOO acts as a bridge to a OpenOffice.org program so a connection
to the running program has to be created first: ::

    >>> import pyoo
    >>> desktop = pyoo.Desktop('localhost', 2002)

Host name and port number used in the example ``('localhost', 2002)``
are default values so they can be omitted.

Connection to a named pipe is also possible: ::

    >>> pyoo.Desktop(pipe='hello')

New spreadsheet document can be created using ``Desktop.create_spreadsheet()``
method or opened using ``Desktop.open_spreadsheet()``: ::

    >>> doc = desktop.create_spreadsheet()
    >>> # doc = desktop.open_spreadsheet("/path/to/spreadsheet.ods")

If the office application is not running in the headless
mode then a new window with Calc program should open now.


Sheets
......

Spreadsheet document is represented by a ``SpreadsheetDocument`` class which
implements basic manipulation with document. All data are in  sheets
which can can be accessed and manipulated via ``SpreadsheetDocument.sheets``
property: ::

    >>> # Access sheet by index or name:
    >>> doc.sheets[0]
    <Sheet: 'Sheet1'>
    >>> doc.sheets['Sheet1']
    <Sheet: 'Sheet1'>

    >>> # Create a new sheet after the first one:
    >>> doc.sheets.create('My Sheet', index=1)
    <Sheet: 'My Sheet'>

    >>> # Copy the created sheet after the second one:
    >>> doc.sheets.copy('My Sheet', 'Copied Sheet', 2)
    <Sheet: 'Copied Sheet'>

    >>> # Delete sheet by index or name:
    >>> del doc.sheets[1]
    >>> del doc.sheets['Copied sheet']

    >>> # Create multiple sheets with same name/prefix
    >>> get_sheet_name = pyoo.NameGenerator()
    >>> doc.sheets.create(get_sheet_name('My sheet'))
    <Sheet: 'My sheet'>
    >>> doc.sheets.create(get_sheet_name('My sheet'))
    <Sheet: 'My sheet 2'>

Cells can be accessed using index notation from a sheet: ::

    >>> # Get sheet:
    >>> sheet = doc.sheets[0]

    >>> # Get cell address and set cell values:
    >>> str(sheet[0,0].address)
   '$A$1'
    >>> sheet[0,0].value = 1
    >>> str(sheet[0,1].address)
    '$B$1'
    >>> sheet[0,1].value = 2

    >>> # Set cell formula and get value:
    >>> sheet[0,2].formula = '=$A$1+$B$1'
    >>> sheet[0,2].value
    3.0

All the changes should be visible in the opened document.

Every operation with a cell takes some time so setting all values separately
is very ineffective. For this reason operations with whole cell ranges
are implemented: ::

    >>> # Tabular (two dimensional) cell range:
    >>> sheet[1:3,0:2].values = [[3, 4], [5, 6]]

    >>> # Row (one dimensional) cell range:
    >>> sheet[3, 0:2].formulas = ['=$A$1+$A$2+$A$3', '=$B$1+$B$2+$B$3']
    >>> sheet[3, 0:2].values
    (9.0, 12.0)

    >>> # Column (one dimensional) cell range:
    >>> sheet[1:4,2].formulas = ['=$A$2+$B$2', '=$A$3+$B3', '=$A$4+$B$4']
    >>> sheet[1:4,2].values
    (7.0, 11.0, 21.0)


Formating
.........

Miscellaneous attributes can be set to cells, cell ranges and sheets
(they all inherit a ``CellRange`` class). Also note that cell ranges
support many indexing options: ::

    >>> # Get cell range with all data
    >>> cells = sheet[:4,:3]

    >>> # Font and text properties:
    >>> cells.font_size = 20
    >>> cells[3, :].font_weight = pyoo.FONT_WEIGHT_BOLD
    >>> cells[:, 2].text_align = pyoo.TEXT_ALIGN_LEFT
    >>> cells[-1,-1].underline = pyoo.UNDERLINE_DOUBLE

    >>> # Colors:
    >>> cells[:3,:2].text_color = 0xFF0000                 # 0xRRGGBB
    >>> cells[:-1,:-1].background_color = 0x0000FF         # 0xRRGGBB

    >>> # Borders
    >>> cells[:,:].border_width = 100
    >>> cells[-4:-1,-3:-1].inner_border_width = 50

Number format can be also set but it is locale dependent: ::

    >>> locale = doc.get_locale('en', 'us')
    >>> sheet.number_format = locale.format(pyoo.FORMAT_PERCENT_INT)


Charts
......

Charts can be created: ::

    >>> chart = sheet.charts.create('My Chart', sheet[5:10, 0:5], sheet[:4,:3])

The first argument is a chart name, the second argument specifies
chart position and the third one contains address of source data
(it can be also a list or tuple). If optional ``row_header`` or
``col_header`` keyword arguments are set to ``True`` then labels
will be read from first row or column.

Existing charts can be accessed either by an index or a name: ::

    >>> sheet.charts[0].name
    u'My Chart'
    >>> sheet.charts['My Chart'].name
    u'My Chart'


Chart instances are generally only a container for diagrams which specify
how are data rendered. Diagram can be replaced by another type while chart
stays same. ::

    >>> chart.diagram.__class__
    <class 'pyoo.BarDiagram'>
    >>> diagram = chart.change_type(pyoo.LineDiagram)
    >>> diagram.__class__
    <class 'pyoo.LineDiagram'>

Diagram instance can be used for accessing and setting of
miscellanous properties. ::

    >>> # Set axis label
    >>> diagram.y_axis.title = "Primary axis"

    >>> # Axis can use a logarithmic scale
    >>> diagram.y_axis.logarithmic = True

    >>> # Secondary axis can be shown.
    >>> diagram.secondary_y_axis.visible = True

    >>> # All axes have same attributes.
    >>> diagram.secondary_y_axis.title = "Secondary axis"

    >>> # Change color of one of series (lines, bars,...)
    >>> diagram.series[0].fill_color = 0x000000

    >>> # And bind it to secondary axis
    >>> diagram.series[0].axis = pyoo.AXIS_SECONDARY


Saving documents
................

Spreadsheet documents can be saved using save method: ::

    >>> doc.save('example.xlsx', pyoo.FILTER_EXCEL_2007)
    >>> # doc.save()

And finally do not forget to close the document: ::

    >>> doc.close()


Testing
-------

Automated integration tests cover most of the code.

The test suite assumes that OpenOffice or LibreOffice is running and
it is listening on localhost port 2002.

All tests are in the ``test.py`` file: ::

    $ python test.py


License
-------

This library is released under the MIT license. Seet the ``LICENSE`` file.
Copyright (c) 2016 Seznam.cz, a.s.
