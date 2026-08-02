# vybe-test: python/list_comprehension/list_comp_nested_filter
# origin: languages/python/tests/python/test_list_comprehension.rs

[x for x in [1, 2, 3, 4] if x > 1 if x < 4]
