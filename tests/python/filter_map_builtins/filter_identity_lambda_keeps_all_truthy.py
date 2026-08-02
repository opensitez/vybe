# vybe-test: python/filter_map_builtins/filter_identity_lambda_keeps_all_truthy
# origin: languages/python/tests/python/test_filter_map_builtins.rs

list(filter(lambda x: x, [1, 2, 3]))
