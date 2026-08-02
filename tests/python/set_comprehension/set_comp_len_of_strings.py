# vybe-test: python/set_comprehension/set_comp_len_of_strings
# origin: languages/python/tests/python/test_set_comprehension.rs

sorted({len(s) for s in ['a', 'bb', 'a']})
