# vybe-test: python/py_scope_namespaces/test_py_comprehension_scope_isolation_py3
# origin: languages/python/tests/python/test_py_scope_namespaces.rs

x = "outside"
lst = [x for x in range(3)]
print(lst)
print(x)  # x in outer scope not leaked in Python 3!
