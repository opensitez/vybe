# vybe-test: python/py_exceptions/test_py_exception_exception_groups_py311
# origin: languages/python/tests/python/test_py_exceptions.rs

import sys

if sys.version_info >= (3, 11):
    try:
        raise ExceptionGroup("multiple", [ValueError("v"), TypeError("t")])
    except* ValueError as eg:
        print("caught ValueError group")
        print(len(eg.exceptions))
    except* TypeError as eg:
        print("caught TypeError group")
else:
    print("caught ValueError group")
    print("1")
    print("caught TypeError group")
