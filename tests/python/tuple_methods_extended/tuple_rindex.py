# vybe-test: python/tuple_methods_extended/tuple_rindex
# origin: languages/python/tests/python/test_tuple_methods_extended.rs
# `tuple` has no `rindex` — only `list` does not either; it is a STR method.
t = (1, 2, 3, 2)
try:
    t.rindex(2)
except AttributeError:
    pass
