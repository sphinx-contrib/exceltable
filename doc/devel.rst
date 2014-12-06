.. _devel:

===========
Development
===========
This section contains some information about module development - in a case you want to contribute to it.
Which is welcome, btw.

.. contents::
   :local:

.. index:: Building


Getting started
===============
To get started with the development, follow the steps:

- Install development-time dependencies::

    pip install nose


Building
========
Build project package with command::

     python setup.py bdist_egg sdist


.. index:: Testing

.. _testing:

Testing
=======
Project uses `nose`_ for unit testing, `coverage`_ for testing coverage reporting and `tox`_
for compliance testing. To execute the tests, run:

- Unittests: ``python setup.py test``
- Compliance: ``tox``

Project repository comes with ready-made configuration for both of the tools, which are used
automatically.


API
====
This section provides some further information about internals of the module:

.. automodule:: sphinxcontrib.exceltable

.. autoclass:: sphinxcontrib.exceltable.ExcelTableDirective

.. autoclass:: sphinxcontrib.exceltable.ExcelTable

.. automethod:: sphinxcontrib.exceltable.ExcelTable.create_table


Licensing
=========
The software is licensed with liberal MIT license, making it suitable for both
commercial and open source usage:

    .. include:: ../LICENSE

