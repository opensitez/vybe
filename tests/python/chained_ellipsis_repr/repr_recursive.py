# vybe-test: python/chained_ellipsis_repr/repr_recursive
# origin: languages/python/tests/python/test_chained_ellipsis_repr.rs
# vybe-test-mode: compile

a = []
a.append(a)
repr(a)
