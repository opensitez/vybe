# vybe-test: python/py_enum_flags_strenum/test_py_str_enum_string_interop
# origin: languages/python/tests/python/test_py_enum_flags_strenum.rs

import sys

if sys.version_info >= (3, 11):
    from enum import StrEnum
    class HttpMethod(StrEnum):
        GET = "GET"
        POST = "POST"
    print(HttpMethod.GET == "GET")
    print(isinstance(HttpMethod.POST, str))
else:
    print("True")
    print("True")
