# vybe-test: python/set_comprehension/set_comp_exclude_spaces
# origin: languages/python/tests/python/test_set_comprehension.rs

sorted({c for c in 'a b' if c != ' '})
