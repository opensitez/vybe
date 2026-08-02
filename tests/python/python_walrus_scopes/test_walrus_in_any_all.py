# vybe-test: python/python_walrus_scopes/test_walrus_in_any_all
# origin: languages/python/tests/python/test_python_walrus_scopes.rs

data = [0, 0, 3, 0, 5]
found = any((first_nonzero := x) for x in data if x != 0)
print(found)
print(first_nonzero)
