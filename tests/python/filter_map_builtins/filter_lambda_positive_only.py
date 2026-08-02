# vybe-test: python/filter_map_builtins/filter_lambda_positive_only
# origin: languages/python/tests/python/test_filter_map_builtins.rs

list(filter(lambda x: x > 0, [-1, 0, 2, 3]))
