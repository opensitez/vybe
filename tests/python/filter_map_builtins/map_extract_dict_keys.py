# vybe-test: python/filter_map_builtins/map_extract_dict_keys
# origin: languages/python/tests/python/test_filter_map_builtins.rs

list(map(lambda kv: kv[0], {'x': 1, 'y': 2}.items()))
