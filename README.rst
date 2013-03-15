
=====================================================
PyOO - Pythonic interface to OpenOffice.org API (UNO)
=====================================================

PyOO allows to control a running OpenOffice.org_ program and
use it for reading and writing spreadsheet documents.
The library can be used for creation of miscellaneous exports
to spreadsheet many document formats -- including Microsoft
Excel 97 (.xls) and Microsoft Excel 2007 (.xlsx).

The main advantage of PyOO library is that it can use almost any
functionality implemented in OpenOffice.org so it does not need
to reinvent the wheel as many libraries do. On the other
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

Accessing documents
...................

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


Sheets
......

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

The first argument is a chart name, the second argument specifies chart
position a the third one contains address of source data (it can be also
a list or tuple). If optional ``row_header`` or ``col_header`` keyword
arguments are set to ``True`` then labels will be read from first row
or column.

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

And finally never forget to close the document: ::

    >>> doc.close()



Testing
-------

Automated integration tests cover most of the code.

Many new features were added to ``unittest`` module in Python 2.7 and
tests for PyOO library use some of them. If you are using older version
of Python please install ``unittest2`` library which back-ports these
features (for example install Debian package `python-unittest2`).

The test suite assumes that OpenOffice.org program is running and
it is listening on localhost port 2002.

All tests are in the ``test.py`` file: ::

    $ python test.py


Release
-------

Make sure that you have the latest version: ::

    $ git pull

Finalize changelog: ::

    $ dch -r

Commit the changes: ::

    $ git add debian/changelog
    $ git commit -m "Release 0.1 version. Refs #439."

Create a debian package: ::

    $ debuild -uc -us -I -b

Upload package to repository and update it: ::

    $ scp ../szn-python-pyoo_X.Y_all.deb debian.kancelar.seznam.cz:/deb/squeeze-u/lib/
    $ ssh debian.kancelar.seznam.cz "makepkginc squeeze-u"


Clean temporary files:

    $ debclean

Tag released version: ::

    $ git tag -a 'vX.Y' -m 'Tag vX.Y version'

Increment current version: ::

  $ dch -i
  $ vi setup.py
  $ git add debian/changelog setup.py
  $ git commit -m "Increment version number. Refs #439."

Push all changes to CML: ::

    $ git push
    $ git push --tags
