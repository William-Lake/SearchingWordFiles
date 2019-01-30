=================
search_word_files
=================

A small Python utility program for searching Word files for some given text.


* Free software: MIT license
* Documentation: https://search-word-files.readthedocs.io.

Dependencies
--------

If you'd only like to use this module, you'll only need the following dependencies:

- fs
- pywin32
- PySimpleGui

Which can be installed via the requirements.txt file: `pip install -r requirements.txt`

If you'd like to develop the module as well, you'll need the requirements_dev.txt file: `pip install -r requirements_dev.txt`

Usage
--------

#. Ensure you have python3 and the dependencies installed.
#. Open a terminal in the same directory as search_word_docs.py
#. Execute `python3 search_word_docs.py`
#. Select a directory to search for word files.
#. Provide a search term.
#. Select/UnSelect the Recursive Searching Textbox
#. Click Search

Updates are provided throughout the process, when finished the results will be provided in the bottom-most text box.

Credits
-------

This package was created with Cookiecutter_ and the `audreyr/cookiecutter-pypackage`_ project template.

.. _Cookiecutter: https://github.com/audreyr/cookiecutter
.. _`audreyr/cookiecutter-pypackage`: https://github.com/audreyr/cookiecutter-pypackage
