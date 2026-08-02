# vybe-test: python/py_enum/test_py_enum_auto_values
# origin: languages/python/tests/python/test_py_enum.rs

from enum import Enum, auto

class Direction(Enum):
    NORTH = auto()
    SOUTH = auto()
    EAST = auto()
    WEST = auto()

print([d.value for d in Direction])
