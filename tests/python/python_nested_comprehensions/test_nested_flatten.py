# vybe-test: python/python_nested_comprehensions/test_nested_flatten
# origin: languages/python/tests/python/test_python_nested_comprehensions.rs

nested = [[1, 2, 3], [4, 5], [6, 7, 8, 9]]
flat = [x for sublist in nested for x in sublist]
print(flat)
