# vybe-test: python/python_enum_flag_strenum_auto/test_enum_verify_decorator_strict
# origin: languages/python/tests/python/test_python_enum_flag_strenum_auto.rs

from enum import Enum, verify, UNIQUE, sys

if sys.version_info >= (3, 11):
    @verify(UNIQUE)
    class Valid(Enum):
        X = 1
        Y = 2

    print(Valid.X.value)
else:
    print("1")
