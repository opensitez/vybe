# vybe-test: python/python_dataclass_kwonly_slots/test_dataclass_kw_only_flag
# origin: languages/python/tests/python/test_python_dataclass_kwonly_slots.rs

from dataclasses import dataclass, KW_ONLY
import sys

if sys.version_info >= (3, 10):
    @dataclass
    class Point:
        x: float
        _: KW_ONLY
        y: float

    p = Point(1.0, y=2.0)
    print(p.x, p.y)
    try:
        Point(1.0, 2.0)
    except TypeError:
        print("TypeError")
else:
    print("1.0 2.0\nTypeError")
