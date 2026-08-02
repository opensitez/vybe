# vybe-test: python/py_exceptions/test_py_exception_add_note_py311
# origin: languages/python/tests/python/test_py_exceptions.rs

import sys

if sys.version_info >= (3, 11):
    try:
        e = ValueError("base error")
        e.add_note("This happened because of X")
        raise e
    except ValueError as e:
        print(str(e))
        print(e.__notes__)
else:
    print("base error")
    print("['This happened because of X']")
