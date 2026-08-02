# vybe-test: python/python_nested_comprehensions/test_nested_dict_comprehension
# origin: languages/python/tests/python/test_python_nested_comprehensions.rs

keys = ["a", "b", "c"]
vals = [1, 2, 3]
d = {k: v for k, v in zip(keys, vals)}
print(d)
