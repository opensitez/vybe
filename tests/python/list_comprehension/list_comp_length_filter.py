# vybe-test: python/list_comprehension/list_comp_length_filter
# origin: languages/python/tests/python/test_list_comprehension.rs

[w for w in ['hi', 'hey', 'yo'] if len(w) == 2]
