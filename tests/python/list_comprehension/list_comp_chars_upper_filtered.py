# vybe-test: python/list_comprehension/list_comp_chars_upper_filtered
# origin: languages/python/tests/python/test_list_comprehension.rs

[c.upper() for c in 'ab' if c != 'b']
