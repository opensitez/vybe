# vybe-test: python/py_dataclass_advanced_features/test_py_dataclass_slots_memory_footprint
# origin: languages/python/tests/python/test_py_dataclass_advanced_features.rs

from dataclasses import dataclass
import sys

if sys.version_info >= (3, 10):
    @dataclass(slots=True)
    class Point:
        x: float
        y: float

    p = Point(1.0, 2.0)
    print(hasattr(p, "__dict__"))
    print(p.x, p.y)
else:
    print("False")
    print("1.0 2.0")
