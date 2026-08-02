# vybe-test: python/dict_comprehension/dict_comp_filter_none_values_out
# origin: languages/python/tests/python/test_dict_comprehension.rs

{i: v for i, v in enumerate([0, None, 2]) if v is not None}
