# vybe-test: python/filter_map_builtins/filter_dict_items_by_value
# origin: languages/python/tests/python/test_filter_map_builtins.rs

list(filter(lambda kv: kv[1] > 1, {'a': 1, 'b': 2}.items()))
