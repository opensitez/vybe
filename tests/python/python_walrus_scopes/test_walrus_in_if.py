# vybe-test: python/python_walrus_scopes/test_walrus_in_if
# origin: languages/python/tests/python/test_python_walrus_scopes.rs

data = [1, 2, 3, 4, 5]
if (n := len(data)) > 3:
    print(f"long list: {n}")
else:
    print(f"short list: {n}")
