# vybe-test: python/py_closures_hof/test_py_lambda_key_functions
# origin: languages/python/tests/python/test_py_closures_hof.rs

data = [{"name": "Charlie", "age": 35}, {"name": "Alice", "age": 25}, {"name": "Bob", "age": 30}]
by_age = sorted(data, key=lambda p: p["age"])
print([p["name"] for p in by_age])
by_name = sorted(data, key=lambda p: p["name"])
print([p["name"] for p in by_name])
