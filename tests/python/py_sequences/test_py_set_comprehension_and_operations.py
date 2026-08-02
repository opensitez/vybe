# vybe-test: python/py_sequences/test_py_set_comprehension_and_operations
# origin: languages/python/tests/python/test_py_sequences.rs

a = {1, 2, 3, 4}
b = {3, 4, 5, 6}
print(sorted(a | b))    # union
print(sorted(a & b))    # intersection
print(sorted(a - b))    # difference
print(sorted(a ^ b))    # symmetric difference

squares = {x**2 for x in range(6)}
print(sorted(squares))
