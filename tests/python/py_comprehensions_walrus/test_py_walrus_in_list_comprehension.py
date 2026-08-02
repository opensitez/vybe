# vybe-test: python/py_comprehensions_walrus/test_py_walrus_in_list_comprehension
# origin: languages/python/tests/python/test_py_comprehensions_walrus.rs

data = [1, -2, 3, -4, 5]
results = [y for x in data if (y := x ** 2) > 4]
print(results)
