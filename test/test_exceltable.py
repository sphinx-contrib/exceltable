# -*- coding: utf-8 -*-
"""
This module contains the units tests for the sphinxcontrib.exceltable -module.
"""
import doctest
import os

from sphinxcontrib import exceltable
from sphinxcontrib.exceltable import toindex, toname


def test_docs():
    """
    Run doc tests
    """
    doctest.testmod()


def test_directive():
    """
    Minimal verification of the module: run some
    unittests against the module
    """
    path = os.path.abspath(os.path.join(os.path.dirname(__file__), '../doc/example/cartoons.xls'))
    excel_table = exceltable.ExcelTable(open(path, 'r+b'))

    # Load 2x2 part from 2nd sheet
    quad = excel_table.create_table(sheet=1, fromcell='B2', tocell='C3')
    assert len(quad['headers']) == 0
    assert len(quad['rows']) == 2
    # print quad

    assert len(excel_table.create_table(nheader=1, fromcell='A1', tocell='B4')['rows']) == 3
    assert len(excel_table.create_table(nheader=1, fromcell='A1', tocell='A1')['headers']) == 1
    assert len(excel_table.create_table(nheader=1, fromcell='A1', tocell='A1')['rows']) == 0
    assert len(excel_table.create_table(nheader=1, fromcell='B2', tocell='A1')['rows']) == 0

    assert toindex(*toname(3, 3)) == (3, 3)
    assert toindex('Z', 3) == (25, 2)
    assert toindex('AA', 4) == (26, 3)
    assert toindex('AB', 5) == (27, 4)

    # Test max value
    assert len(excel_table.create_table(nheader=0, fromcell=None, tocell='F9')['rows']) == 4
