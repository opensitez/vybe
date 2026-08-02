# vybe-test: python/py_collections/test_py_collections_deque_maxlen
# origin: languages/python/tests/python/test_py_collections.rs

from collections import deque

d = deque(maxlen=3)
for i in range(6):
    d.append(i)
print(list(d))  # keeps only last 3
d.appendleft(99)
print(list(d))  # oldest drops off other end
