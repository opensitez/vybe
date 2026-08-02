# vybe-test: python/python_walrus_scopes/test_walrus_in_comprehension
# origin: languages/python/tests/python/test_python_walrus_scopes.rs

vals = [1, -2, 3, -4, 5]
positive = [y for x in vals if (y := x * 2) > 0]
print(positive)
