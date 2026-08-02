# vybe-test: python/list_comprehension/list_comp_filter_none_values
# origin: languages/python/tests/python/test_list_comprehension.rs

[x for x in [0, 1, None, 2] if x is not None]
