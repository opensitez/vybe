# vybe-test: python/py_comprehensions_walrus/test_py_conditional_expression_ternary
# origin: languages/python/tests/python/test_py_comprehensions_walrus.rs

x = 10
label = "positive" if x > 0 else "non-positive"
print(label)

values = [abs(x) if x >= 0 else -x for x in range(-3, 4)]
print(values)

nested = "big" if x > 100 else "medium" if x > 10 else "small"
print(nested)
