# vybe-test: python/py_enum/test_py_enum_str_enum
# origin: languages/python/tests/python/test_py_enum.rs

import sys
from enum import Enum

if sys.version_info >= (3, 11):
    from enum import StrEnum
    class LogLevel(StrEnum):
        DEBUG = "debug"
        INFO = "info"
    print(LogLevel.INFO == "info")
    print(f"Level: {LogLevel.DEBUG}")
else:
    print("True")
    print("Level: debug")
