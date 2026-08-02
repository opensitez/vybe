# vybe-test: python/py_builtins_adv/test_py_builtins_enumerate_with_start
# origin: languages/python/tests/python/test_py_builtins_adv.rs

items = ["apple", "banana", "cherry"]
for idx, item in enumerate(items, start=1):
    print(f"{idx}: {item}")
