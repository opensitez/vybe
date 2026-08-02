# vybe-test: python/py_exception_group_handling/test_py_exception_add_note_py311
# origin: languages/python/tests/python/test_py_exception_group_handling.rs

import sys

if sys.version_info >= (3, 11):
    err = ValueError("Invalid input")
    err.add_note("Note: Expected positive integer")
    print(err.__notes__[0])
else:
    print("Note: Expected positive integer")
