# vybe-test: python/python_nested_comprehensions/test_nested_set_comprehension
# origin: languages/python/tests/python/test_python_nested_comprehensions.rs

nums = [1, 2, 3, 1, 2, 4]
unique_squares = {x ** 2 for x in nums}
print(sorted(unique_squares))
