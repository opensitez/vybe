# vybe-test: python/python_nested_comprehensions/test_nested_generator_expression
# origin: languages/python/tests/python/test_python_nested_comprehensions.rs

total = sum(x * y for x in range(1, 4) for y in range(1, 4) if x == y)
print(total)  # 1*1 + 2*2 + 3*3 = 14
