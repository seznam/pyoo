
=============================================================
PyOO - Pythonic interface to OpenOffice.org API known as UNO.
=============================================================

PyOO allows to control a running OpenOffice.org_ program and
use it for reading and writing spreadsheet documents.
The library can be used for creation of miscellaneous exports
to spreadsheet many document formats -- including Microsoft
Excel 97 (.xls) and Microsoft Excel 2007 (.xlsx).

The main advantage of PyOO library is that it can use almost any
functionality implemented in OpenOffice.org so it does not need
to reinvent the wheel as many libraries do. On the other other
side it requires a running instance of an OpenOffice.org program
which means significant overhead.

PyOO uses an UNO_ interface via Python-UNO_ bridge. UNO is a
standard interface which allows to control running OpenOffice program.
Python-UNO makes this interface available in Python scripts.

Direct usage of UNO API via Python-UNO can be quite complicated
and even simple tasks require a lot of code. Another problem
is that many UNO calls are slow and should be avoided when it
is possible.

PyOO wraps robust Python-UNO bridge to simple and Pythonic
interface. Under the hood it implements miscellaneous
optimizations which can prevent unnecessary expensive UNO
calls.


.. _OpenOffice.org: http://www.openoffice.org/
.. _UNO: http://www.openoffice.org/api/docs/common/ref/com/sun/star/module-ix.html
.. _Python-UNO: http://www.openoffice.org/udk/python/python-bridge.html


Installation
------------

PyOO library can be installed using standard setup.py script: ::

    $ python setup.py install

The only dependecy is the Python-UNO library (imported as a ``uno`` module).
If OpenOffice.org is installed in the system then Python-UNO should be
already present It can be installed as a ``python-uno`` Debian package.


Starting OpenOffice.org
-----------------------


PyOO requires a running OpenOffice.org instance which it can
connect to. On Debian you can start the program from a command
line by command similar to: ::

    $ openoffice.org -accept="socket,host=localhost,port=2002;urp;StarOffice.ServiceManager" -norestore -nofirstwizard -nologo -invisible # -headless

If you plan to connect remotely replace ``localhost`` by your IP address
or ``0.0.0.0``. Because of the other options no user interface should be
displayed until first document is opened. If the `-headless` option is
used then no user interface is visible even when a document is opened.

On Ubunut use following command: ::

    $ soffice --accept="socket,host=localhost,port=2002;urp;" --norestore --nologo --nodefault # --headless



Usage
-----

PyOO acts as a bridge to a OpenOffice.org program so a connection
to the running program has to be created first: ::

    >>> import pyoo
    >>> desktop = pyoo.Desktop('localhost', 2002)

Host name and port number used in the example ``('localhost', 2002)``
are default values so they can be omitted.

New spreadsheet document can be created using ``Desktop.create_spreadsheet()``
method or opened using ``Desktop.open_spreadsheet()``: ::

    >>> doc = desktop.create_spreadsheet()
    >>> # doc = desktop.open_spreadsheet("/path/to/spreadsheet.ods")

If OpenOffice.org program is not running in a headless mode then
a new window with Calc program should be opened.

Spreasheet document is represented by a ``SpreadsheetDocument`` class which
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

And finally save and close the document: ::

    >>> doc.save('example.xlsx', pyoo.FILTER_EXCEL_2007)
    >>> doc.close()


Testing
-------

The test suite requires OpenOffice.org running and listening
on localhost port 2002. All tests are in the ``test.py`` file: ::

    $ python test.py


