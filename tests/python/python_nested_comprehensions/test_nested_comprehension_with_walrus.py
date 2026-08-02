# vybe-test: python/python_nested_comprehensions/test_nested_comprehension_with_walrus
# origin: languages/python/tests/python/test_python_nested_comprehensions.rs

data = [1, -2, 3, -4, 5]
pos = [y for x in data if (y := abs(x)) > 2]
print(pos)
