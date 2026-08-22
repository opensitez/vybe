# vybe-test: python/comprehensions/list_comp_matrix
# origin: languages/python/tests/python/test_comprehensions.rs

matrix = [[i*j for j in range(3)] for i in range(3)]
