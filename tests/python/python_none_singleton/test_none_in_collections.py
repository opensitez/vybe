# vybe-test: python/python_none_singleton/test_none_in_collections
# origin: languages/python/tests/python/test_python_none_singleton.rs

lst = [1, None, 2, None, 3]
count = lst.count(None)
print(count)
non_none = [x for x in lst if x is not None]
print(non_none)
