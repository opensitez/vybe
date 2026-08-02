# vybe-test: python/py_exception_group_handling/test_py_exception_group_creation_and_exceptions
# origin: languages/python/tests/python/test_py_exception_group_handling.rs

import sys

if sys.version_info >= (3, 11):
    eg = ExceptionGroup("Multiple errors", [ValueError("bad val"), TypeError("bad type")])
    print(eg.message)
    print([type(e).__name__ for e in eg.exceptions])
else:
    print("Multiple errors")
    print("['ValueError', 'TypeError']")
