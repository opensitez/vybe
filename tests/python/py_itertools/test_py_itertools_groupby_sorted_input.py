# vybe-test: python/py_itertools/test_py_itertools_groupby_sorted_input
# origin: languages/python/tests/python/test_py_itertools.rs

import itertools

words = ["apple", "ant", "bear", "bat", "cat"]
for key, group in itertools.groupby(sorted(words), key=lambda w: w[0]):
    print(f"{key}: {list(group)}")
