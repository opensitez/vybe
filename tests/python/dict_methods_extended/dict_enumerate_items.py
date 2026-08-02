# vybe-test: python/dict_methods_extended/dict_enumerate_items
# origin: languages/python/tests/python/test_dict_methods_extended.rs

d = {'x': 1, 'y': 2}
print([i for i, _ in enumerate(d)])
