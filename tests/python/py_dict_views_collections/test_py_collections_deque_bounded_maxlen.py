# vybe-test: python/py_dict_views_collections/test_py_collections_deque_bounded_maxlen
# origin: languages/python/tests/python/test_py_dict_views_collections.rs

from collections import deque

dq = deque(maxlen=3)
for i in range(5):
    dq.append(i)

print(list(dq))
dq.appendleft(99)
print(list(dq))
