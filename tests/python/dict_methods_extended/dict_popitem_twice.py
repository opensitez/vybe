# vybe-test: python/dict_methods_extended/dict_popitem_twice
# origin: languages/python/tests/python/test_dict_methods_extended.rs
# vybe-test-mode: compile

d = {'a': 1, 'b': 2}
d.popitem()
d.popitem()
