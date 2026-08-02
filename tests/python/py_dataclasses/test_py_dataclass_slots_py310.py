# vybe-test: python/py_dataclasses/test_py_dataclass_slots_py310
# origin: languages/python/tests/python/test_py_dataclasses.rs

from dataclasses import dataclass
import sys

if sys.version_info >= (3, 10):
    @dataclass(slots=True)
    class Slotted:
        x: int
        y: int
    s = Slotted(1, 2)
    print(s.x + s.y)
    print(hasattr(s, '__dict__'))
else:
    print("3")
    print("False")
