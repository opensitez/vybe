# vybe-test: python/python_nested_comprehensions/test_nested_with_filter
# origin: languages/python/tests/python/test_python_nested_comprehensions.rs

pairs = [(x, y) for x in range(4) for y in range(4) if x != y and x + y == 3]
print(sorted(pairs))
