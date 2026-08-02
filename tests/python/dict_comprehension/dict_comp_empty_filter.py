# vybe-test: python/dict_comprehension/dict_comp_empty_filter
# origin: languages/python/tests/python/test_dict_comprehension.rs

{x: x for x in range(3) if x > 10}
