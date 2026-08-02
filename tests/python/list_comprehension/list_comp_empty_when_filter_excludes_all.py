# vybe-test: python/list_comprehension/list_comp_empty_when_filter_excludes_all
# origin: languages/python/tests/python/test_list_comprehension.rs

[x for x in range(3) if x > 10]
