# vybe-test: python/python_tuple_operations/test_tuple_immutable
# origin: languages/python/tests/python/test_python_tuple_operations.rs

t = (1, 2, 3)
try:
    t[0] = 99
    print("mutable")
except TypeError:
    print("immutable")
