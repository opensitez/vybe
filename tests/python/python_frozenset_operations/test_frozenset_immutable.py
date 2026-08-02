# vybe-test: python/python_frozenset_operations/test_frozenset_immutable
# origin: languages/python/tests/python/test_python_frozenset_operations.rs

fs = frozenset([1, 2])
try:
    fs.add(3)
    print("no_error")
except AttributeError:
    print("AttributeError")
