# vybe-test: python/py_comprehensions_walrus/test_py_nested_list_comprehension_flatten
# origin: languages/python/tests/python/test_py_comprehensions_walrus.rs

matrix = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
flat = [x for row in matrix for x in row]
print(flat)

# Only even values
evens = [x for row in matrix for x in row if x % 2 == 0]
print(evens)
