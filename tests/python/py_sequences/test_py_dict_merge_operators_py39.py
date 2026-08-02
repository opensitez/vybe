# vybe-test: python/py_sequences/test_py_dict_merge_operators_py39
# origin: languages/python/tests/python/test_py_sequences.rs

import sys

d1 = {"a": 1, "b": 2}
d2 = {"b": 99, "c": 3}

if sys.version_info >= (3, 9):
    merged = d1 | d2
    print(merged)
    d1 |= d2
    print(d1)
else:
    print({**d1, **d2})
    d1.update(d2)
    print(d1)
