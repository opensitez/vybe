# vybe-test: python/filter_map_builtins/filter_on_tuple_of_mixed_types
# origin: languages/python/tests/python/test_filter_map_builtins.rs

list(filter(lambda x: isinstance(x, int), (1, 'a', 2)))
