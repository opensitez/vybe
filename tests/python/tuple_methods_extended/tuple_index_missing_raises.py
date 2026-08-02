# vybe-test: python/tuple_methods_extended/tuple_index_missing_raises
# origin: languages/python/tests/python/test_tuple_methods_extended.rs
# vybe-test-mode: compile

t = (1, 2)
try:
    t.index(9)
except ValueError:
    pass
