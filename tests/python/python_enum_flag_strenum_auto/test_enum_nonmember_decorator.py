# vybe-test: python/python_enum_flag_strenum_auto/test_enum_nonmember_decorator
# origin: languages/python/tests/python/test_python_enum_flag_strenum_auto.rs

from enum import Enum
import sys

if sys.version_info >= (3, 11):
    from enum import nonmember
    class Config(Enum):
        HOST = "localhost"
        helper = nonmember(lambda: "helper_func")

    print(Config.HOST.value)
    print(Config.helper())
else:
    print("localhost\nhelper_func")
