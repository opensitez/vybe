# vybe-test: python/python_enum_flag_strenum_auto/test_enum_iteration_yields_members
# origin: languages/python/tests/python/test_python_enum_flag_strenum_auto.rs

from enum import Enum

class Season(Enum):
    SPRING = 1
    SUMMER = 2
    AUTUMN = 3
    WINTER = 4

names = [s.name for s in Season]
print(names)
