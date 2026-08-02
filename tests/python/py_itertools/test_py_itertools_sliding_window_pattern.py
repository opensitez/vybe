# vybe-test: python/py_itertools/test_py_itertools_sliding_window_pattern
# origin: languages/python/tests/python/test_py_itertools.rs

import itertools, collections

def sliding_window(iterable, n):
    it = iter(iterable)
    window = collections.deque(itertools.islice(it, n), maxlen=n)
    if len(window) == n:
        yield tuple(window)
    for x in it:
        window.append(x)
        yield tuple(window)

print(list(sliding_window([1, 2, 3, 4, 5], 3)))
