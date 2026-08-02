# vybe-test: python/python_walrus_scopes/test_walrus_escapes_comprehension_scope
# origin: languages/python/tests/python/test_python_walrus_scopes.rs

last = None
data = [1, 2, 3, 4, 5]
result = [last := x for x in data]
print(last)
print(result)
