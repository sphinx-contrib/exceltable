Module ``sphinxcontrib.exceltable`` is an extension for Sphinx_, that adds support for including Excel spreadsheets, or part
of them, into Sphinx document. See documentation for further information. It works with Python >=3.7

Installation::

    mkdir my-docs
    cd my-docs/

    # Install dependencies
    python3 -v venv
    source vevn/bin/activate
    pip3 install sphinx sphinxcontrib-exceltable

    # Alternatively, install pre-release
    # pip3 install sphinxcontrib-exceltable --pre

    # Create simple docs
    sphinx-quickstart

Configuration:

Enable the extension by adding ::


  extensions = [
    'sphinxcontrib.exceltable'
    # ...other extensions
  ]

Usage::

  My document
  ===========
  The contents of the setup script:

  .. exceltable:: Table caption
     :file: path/to/document.xls
     :header: 1
     :selection: A1:B3

Read complete documentation: http://pythonhosted.org/sphinxcontrib-exceltable/
Report issues: https://github.com/sphinx-contrib/exceltable/issues

Development::

  # Create virtual environment
  python3 -m venv venv3
  source venv3/bin/activate

  # Install dependencies
  python3 -m pip install --upgrade pip
  pip3 install -r requirements.txt

  # Run tests
  PYTHONPATH=$(pwd)/src python3 -m pytest

  # Run
  python3 -m tox

.. _Sphinx: https://www.sphinx-doc.org/
