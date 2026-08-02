# vybe-test: python/filter_map_builtins/filter_lambda_none_explicitly_false
# origin: languages/python/tests/python/test_filter_map_builtins.rs

list(filter(lambda x: x is not None, [None, 1, None, 2]))
