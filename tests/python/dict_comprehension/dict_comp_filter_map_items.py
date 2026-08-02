# vybe-test: python/dict_comprehension/dict_comp_filter_map_items
# origin: languages/python/tests/python/test_dict_comprehension.rs

{k: v for k, v in {'a': 1, 'b': 0}.items() if v}
