# vybe-test: python/list_comprehension/list_comp_nested_flattens_pairs
# origin: languages/python/tests/python/test_list_comprehension.rs

[b for a in [1, 2] for b in [a, a + 10]]
