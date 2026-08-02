# vybe-test: python/set_comprehension/set_comp_chars_isdigit
# origin: languages/python/tests/python/test_set_comprehension.rs

sorted({c for c in 'a1b2' if c.isdigit()})
