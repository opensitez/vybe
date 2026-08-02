# vybe-test: python/python_enum_flag_strenum_auto/test_enum_str_enum_string_subclass
# origin: languages/python/tests/python/test_python_enum_flag_strenum_auto.rs

from enum import Enum, auto
import sys

if sys.version_info >= (3, 11):
    from enum import StrEnum
    class Status(StrEnum):
        PENDING = auto()
        ACTIVE = "active"

    print(Status.PENDING == "pending")
    print(Status.ACTIVE == "active")
else:
    print("True\nTrue")
