# vybe-test: python/python_nested_comprehensions/test_nested_conditional_comprehension
# origin: languages/python/tests/python/test_python_nested_comprehensions.rs

result = ["even" if x % 2 == 0 else "odd" for x in range(6)]
print(result)
