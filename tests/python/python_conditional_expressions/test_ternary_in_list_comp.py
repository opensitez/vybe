# vybe-test: python/python_conditional_expressions/test_ternary_in_list_comp
# origin: languages/python/tests/python/test_python_conditional_expressions.rs

data = [1, -2, 3, -4, 5]
signs = ["pos" if x > 0 else "neg" for x in data]
print(signs)
