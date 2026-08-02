# vybe-test: python/set_comprehension/set_comp_string_stripped
# origin: languages/python/tests/python/test_set_comprehension.rs

sorted({s.strip() for s in [' a', 'b ']})
