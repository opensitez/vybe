# vybe-test: python/python_enum_advanced/test_enum_iteration
# origin: languages/python/tests/python/test_python_enum_advanced.rs

from enum import Enum

class Day(Enum):
    MON = 1
    TUE = 2
    WED = 3

for d in Day:
    print(d.name)
