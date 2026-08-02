# vybe-test: python/py_numeric_math_semantics/test_py_int_bit_length_bit_count
# origin: languages/python/tests/python/test_py_numeric_math_semantics.rs

import sys

n = 255
print(n.bit_length())
if sys.version_info >= (3, 10):
    print(n.bit_count())
else:
    print(8)
