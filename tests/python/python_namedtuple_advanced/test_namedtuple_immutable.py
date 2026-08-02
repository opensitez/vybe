# vybe-test: python/python_namedtuple_advanced/test_namedtuple_immutable
# origin: languages/python/tests/python/test_python_namedtuple_advanced.rs

from collections import namedtuple
Point = namedtuple('Point', ['x', 'y'])
p = Point(1, 2)
try:
    p.x = 99
    print("mutable")
except AttributeError:
    print("immutable")
