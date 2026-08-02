# vybe-test: python/python_dataclass_kwonly_slots/test_dataclass_match_args_generated
# origin: languages/python/tests/python/test_python_dataclass_kwonly_slots.rs

from dataclasses import dataclass
import sys

if sys.version_info >= (3, 10):
    @dataclass
    class Point:
        x: int
        y: int

    print(Point.__match_args__)
else:
    print("('x', 'y')")
