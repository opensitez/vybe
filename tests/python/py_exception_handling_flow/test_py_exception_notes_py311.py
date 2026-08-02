# vybe-test: python/py_exception_handling_flow/test_py_exception_notes_py311
# origin: languages/python/tests/python/test_py_exception_handling_flow.rs

import sys

if sys.version_info >= (3, 11):
    try:
        err = ValueError("bad value")
        err.add_note("Context info: field X")
        raise err
    except ValueError as e:
        print(e.__notes__[0])
else:
    print("Context info: field X")
