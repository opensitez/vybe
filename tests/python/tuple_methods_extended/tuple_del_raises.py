# vybe-test: python/tuple_methods_extended/tuple_del_raises
# origin: languages/python/tests/python/test_tuple_methods_extended.rs
# vybe-test-mode: compile

t = (1, 2)
try:
    del t[0]
except TypeError:
    pass
