# vybe-test: python/py_collections_counter_deque_chainmap/test_py_counter_total_method_py310
# origin: languages/python/tests/python/test_py_collections_counter_deque_chainmap.rs

import sys
from collections import Counter

c = Counter(a=10, b=20, c=30)
if sys.version_info >= (3, 10):
    print(c.total())
else:
    print(sum(c.values()))
