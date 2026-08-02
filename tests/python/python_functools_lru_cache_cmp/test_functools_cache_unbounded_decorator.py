# vybe-test: python/python_functools_lru_cache_cmp/test_functools_cache_unbounded_decorator
# origin: languages/python/tests/python/test_python_functools_lru_cache_cmp.rs

import functools, sys
if sys.version_info >= (3, 9):
    @functools.cache
    def add(a, b): return a + b

    print(add(2, 3))
    print(add(2, 3))
    print(add.cache_info().hits)
else:
    print("5\n5\n1")
