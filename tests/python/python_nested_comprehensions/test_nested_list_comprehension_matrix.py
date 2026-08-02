# vybe-test: python/python_nested_comprehensions/test_nested_list_comprehension_matrix
# origin: languages/python/tests/python/test_python_nested_comprehensions.rs

matrix = [[i * j for j in range(1, 4)] for i in range(1, 4)]
print(matrix)
