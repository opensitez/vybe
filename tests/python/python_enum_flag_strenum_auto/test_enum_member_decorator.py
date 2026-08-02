# vybe-test: python/python_enum_flag_strenum_auto/test_enum_member_decorator
# origin: languages/python/tests/python/test_python_enum_flag_strenum_auto.rs

from enum import Enum
import sys

if sys.version_info >= (3, 11):
    from enum import member
    class FnEnum(Enum):
        ADD = member(lambda x, y: x + y)

    print(FnEnum.ADD.value(2, 3))
else:
    print("5")
