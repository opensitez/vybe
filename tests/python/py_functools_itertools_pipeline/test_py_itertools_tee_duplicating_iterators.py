# vybe-test: python/py_functools_itertools_pipeline/test_py_itertools_tee_duplicating_iterators
# origin: languages/python/tests/python/test_py_functools_itertools_pipeline.rs

from itertools import tee

gen = (x * x for x in range(5))
it1, it2 = tee(gen, 2)

print(list(it1))
print(list(it2))  # independent copies!
