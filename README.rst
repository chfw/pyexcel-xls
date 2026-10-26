================================================================================
pyexcel-xls - Let you focus on data, instead of xls format
================================================================================

.. image:: https://raw.githubusercontent.com/pyexcel/pyexcel.github.io/master/images/patreon.png
   :target: https://www.patreon.com/chfw

.. image:: https://raw.githubusercontent.com/pyexcel/pyexcel-mobans/master/images/awesome-badge.svg
   :target: https://awesome-python.com/#specific-formats-processing

.. image:: https://codecov.io/gh/pyexcel/pyexcel-xls/branch/master/graph/badge.svg
   :target: https://codecov.io/gh/pyexcel/pyexcel-xls

.. image:: https://badge.fury.io/py/pyexcel-xls.svg
   :target: https://pypi.org/project/pyexcel-xls

.. image:: https://anaconda.org/conda-forge/pyexcel-xls/badges/version.svg
   :target: https://anaconda.org/conda-forge/pyexcel-xls


.. image:: https://pepy.tech/badge/pyexcel-xls/month
   :target: https://pepy.tech/project/pyexcel-xls

.. image:: https://anaconda.org/conda-forge/pyexcel-xls/badges/downloads.svg
   :target: https://anaconda.org/conda-forge/pyexcel-xls

.. image:: https://img.shields.io/gitter/room/gitterHQ/gitter.svg
   :target: https://gitter.im/pyexcel/Lobby

.. image:: https://img.shields.io/static/v1?label=continuous%20templating&message=%E6%A8%A1%E7%89%88%E6%9B%B4%E6%96%B0&color=blue&style=flat-square
    :target: https://moban.readthedocs.io/en/latest/#at-scale-continous-templating-for-open-source-projects

.. image:: https://img.shields.io/static/v1?label=coding%20style&message=black&color=black&style=flat-square
    :target: https://github.com/psf/black

**pyexcel-xls** is a tiny wrapper library to read, manipulate and
write data in xls format and it can read xlsx and xlsm fromat.
You are likely to use it with `pyexcel <https://github.com/pyexcel/pyexcel>`_.

Oct 2021 - Update:
===================

1. v0.7.0 removed the pin on xlrd < 2. If you have xlrd >= 2, this
library will NOT read 'xlsx' format and you need to install pyexcel-xlsx. Othwise,
this library can use xlrd < 2 to read xlsx format for you. So 'xlsx' support
in this library will vary depending on the installed version of xlrd.

2. v0.7.0 can write datetime.timedelta. but when the value is read out,
you will get datetime.datetime. so you as the developer decides what to do with it.

Past news
===========

`detect_merged_cells` allows you to spread the same value among
all merged cells. But be aware that this may slow down its reading
performance.

`skip_hidden_row_and_column` allows you to skip hidden rows
and columns and is defaulted to **True**. It may slow down its reading
performance. And it is only valid for 'xls' files. For 'xlsx' files,
please use pyexcel-xlsx.

Date-system override
====================

By default, dates use the workbook's stored date system. If a workbook has an
incorrect date-system flag, pass ``date_mode=0`` for the 1900 system or
``date_mode=1`` for the 1904 system when reading it. This changes how date serials
are interpreted; it does not change the workbook. ``date_mode=None`` retains the
workbook's date system. For example::

    data = get_data("dates.xls", date_mode=1)

Warning
================================================================================

**xls file cannot contain more than 65,000 rows**. You are risking the reputation
of yourself/your company/
`your country <https://www.bbc.co.uk/news/technology-54423988>`_ if you keep
using xls and are not aware of its row limit.


Support the project
================================================================================

If your company uses pyexcel and its components in a revenue-generating product,
please consider supporting the project on GitHub or
`Patreon <https://www.patreon.com/bePatron?u=5537627>`_. Your financial
support will enable me to dedicate more time to coding, improving documentation,
and creating engaging content.


Known constraints
==================

Fonts, colors and charts are not supported.

Nor to read password protected xls, xlsx and ods files.

Installation
================================================================================


You can install pyexcel-xls via pip:

.. code-block:: bash

    $ pip install pyexcel-xls


or clone it and install it:

.. code-block:: bash

    $ git clone https://github.com/pyexcel/pyexcel-xls.git
    $ cd pyexcel-xls
    $ python setup.py install

Usage
================================================================================

As a standalone library
--------------------------------------------------------------------------------

.. testcode::
   :hide:

    >>> import os
    >>> import sys
    >>> from io import BytesIO
    >>> from collections import OrderedDict


Write to an xls file
********************************************************************************



Here's the sample code to write a dictionary to an xls file:

.. code-block:: python

    >>> from pyexcel_xls import save_data
    >>> data = OrderedDict() # from collections import OrderedDict
    >>> data.update({"Sheet 1": [[1, 2, 3], [4, 5, 6]]})
    >>> data.update({"Sheet 2": [["row 1", "row 2", "row 3"]]})
    >>> save_data("your_file.xls", data)


Read from an xls file
********************************************************************************

Here's the sample code:

.. code-block:: python

    >>> from pyexcel_xls import get_data
    >>> data = get_data("your_file.xls")
    >>> import json
    >>> print(json.dumps(data))
    {"Sheet 1": [[1, 2, 3], [4, 5, 6]], "Sheet 2": [["row 1", "row 2", "row 3"]]}


Write an xls to memory
********************************************************************************

Here's the sample code to write a dictionary to an xls file:

.. code-block:: python

    >>> from pyexcel_xls import save_data
    >>> data = OrderedDict()
    >>> data.update({"Sheet 1": [[1, 2, 3], [4, 5, 6]]})
    >>> data.update({"Sheet 2": [[7, 8, 9], [10, 11, 12]]})
    >>> io = BytesIO()
    >>> save_data(io, data)
    >>> # do something with the io
    >>> # In reality, you might give it to your http response
    >>> # object for downloading




Read from an xls from memory
********************************************************************************

Continue from previous example:

.. code-block:: python

    >>> # This is just an illustration
    >>> # In reality, you might deal with xls file upload
    >>> # where you will read from requests.FILES['YOUR_XLS_FILE']
    >>> data = get_data(io)
    >>> print(json.dumps(data))
    {"Sheet 1": [[1, 2, 3], [4, 5, 6]], "Sheet 2": [[7, 8, 9], [10, 11, 12]]}


Pagination feature
********************************************************************************



Let's assume the following file is a huge xls file:

.. code-block:: python

   >>> huge_data = [
   ...     [1, 21, 31],
   ...     [2, 22, 32],
   ...     [3, 23, 33],
   ...     [4, 24, 34],
   ...     [5, 25, 35],
   ...     [6, 26, 36]
   ... ]
   >>> sheetx = {
   ...     "huge": huge_data
   ... }
   >>> save_data("huge_file.xls", sheetx)

And let's pretend to read partial data:

.. code-block:: python

   >>> partial_data = get_data("huge_file.xls", start_row=2, row_limit=3)
   >>> print(json.dumps(partial_data))
   {"huge": [[3, 23, 33], [4, 24, 34], [5, 25, 35]]}

And you could as well do the same for columns:

.. code-block:: python

   >>> partial_data = get_data("huge_file.xls", start_column=1, column_limit=2)
   >>> print(json.dumps(partial_data))
   {"huge": [[21, 31], [22, 32], [23, 33], [24, 34], [25, 35], [26, 36]]}

Obvious, you could do both at the same time:

.. code-block:: python

   >>> partial_data = get_data("huge_file.xls",
   ...     start_row=2, row_limit=3,
   ...     start_column=1, column_limit=2)
   >>> print(json.dumps(partial_data))
   {"huge": [[23, 33], [24, 34], [25, 35]]}

.. testcode::
   :hide:

   >>> os.unlink("huge_file.xls")


As a pyexcel plugin
--------------------------------------------------------------------------------

No longer, explicit import is needed since pyexcel version 0.2.2. Instead,
this library is auto-loaded. So if you want to read data in xls format,
installing it is enough.


Reading from an xls file
********************************************************************************

Here is the sample code:

.. code-block:: python

    >>> import pyexcel as pe
    >>> sheet = pe.get_book(file_name="your_file.xls")
    >>> sheet
    Sheet 1:
    +---+---+---+
    | 1 | 2 | 3 |
    +---+---+---+
    | 4 | 5 | 6 |
    +---+---+---+
    Sheet 2:
    +-------+-------+-------+
    | row 1 | row 2 | row 3 |
    +-------+-------+-------+


Writing to an xls file
********************************************************************************

Here is the sample code:

.. code-block:: python

    >>> sheet.save_as("another_file.xls")


Reading from a IO instance
********************************************************************************

You got to wrap the binary content with stream to get xls working:

.. code-block:: python

    >>> # This is just an illustration
    >>> # In reality, you might deal with xls file upload
    >>> # where you will read from requests.FILES['YOUR_XLS_FILE']
    >>> xlsfile = "another_file.xls"
    >>> with open(xlsfile, "rb") as f:
    ...     content = f.read()
    ...     r = pe.get_book(file_type="xls", file_content=content)
    ...     print(r)
    ...
    Sheet 1:
    +---+---+---+
    | 1 | 2 | 3 |
    +---+---+---+
    | 4 | 5 | 6 |
    +---+---+---+
    Sheet 2:
    +-------+-------+-------+
    | row 1 | row 2 | row 3 |
    +-------+-------+-------+


Writing to a BytesIO instance
********************************************************************************

You need to pass a BytesIO instance to Writer:

.. code-block:: python

    >>> data = [
    ...     [1, 2, 3],
    ...     [4, 5, 6]
    ... ]
    >>> io = BytesIO()
    >>> sheet = pe.Sheet(data)
    >>> io = sheet.save_to_memory("xls", io)
    >>> # then do something with io
    >>> # In reality, you might give it to your http response
    >>> # object for downloading


License
================================================================================

New BSD License

Developer guide
==================

Development steps for code changes

#. git clone https://github.com/pyexcel/pyexcel-xls.git
#. cd pyexcel-xls

Upgrade your setup tools and pip. They are needed for development and testing only:

#. pip install --upgrade setuptools pip

Then install relevant development requirements:

#. pip install -r rnd_requirements.txt # if such a file exists
#. pip install -r requirements.txt
#. pip install -r tests/requirements.txt

Once you have finished your changes, please provide test case(s), relevant documentation
and update changelog.yml

.. note::

    As to rnd_requirements.txt, usually, it is created when a dependent
    library is not released. Once the dependency is installed
    (will be released), the future
    version of the dependency in the requirements.txt will be valid.


How to test your contribution
--------------------------------------------------------------------------------

Although `nose` and `doctest` are both used in code testing, it is advisable
that unit tests are put in tests. `doctest` is incorporated only to make sure
the code examples in documentation remain valid across different development
releases.

On Linux/Unix systems, please launch your tests like this::

    $ make

On Windows, please issue this command::

    > test.bat


Before you commit
------------------------------

Please run::

    $ make format

so as to beautify your code otherwise your build may fail your unit test.


Known Issues
=============

* If a zero was typed in a DATE formatted field in xls, you will get "01/01/1900".
* If a zero was typed in a TIME formatted field in xls, you will get "00:00:00".

.. testcode::
   :hide:

   >>> import os
   >>> os.unlink("your_file.xls")
   >>> os.unlink("another_file.xls")
