# vybe-test: python/python_scope_closures/test_comprehension_scope_isolated
# origin: languages/python/tests/python/test_python_scope_closures.rs

x = 10
result = [x for x in range(3)]
print(x)  # x unchanged after comprehension
print(result)
