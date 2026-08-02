# vybe-test: python/py_set_frozenset_algebra/test_py_set_comprehension_filtering_transform
# origin: languages/python/tests/python/test_py_set_frozenset_algebra.rs

words = ["Apple", "banana", "Cherry", "APPLE", "Banana"]
normalized = {w.lower() for w in words}
print(sorted(normalized))
