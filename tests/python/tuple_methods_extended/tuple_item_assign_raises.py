# vybe-test: python/tuple_methods_extended/tuple_item_assign_raises
# origin: languages/python/tests/python/test_tuple_methods_extended.rs

t = (1, 2)
try:
    t[0] = 9
except TypeError:
    pass
