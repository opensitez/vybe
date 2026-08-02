# vybe-test: python/py_comprehensions_walrus/test_py_unpacking_in_comprehension_target
# origin: languages/python/tests/python/test_py_comprehensions_walrus.rs

pairs = [("alice", 30), ("bob", 25), ("carol", 35)]
names = [name for name, age in pairs if age > 25]
print(names)

age_map = {name: age for name, age in pairs}
print(age_map)
