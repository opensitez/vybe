# vybe-test: python/filter_map_builtins/filter_lambda_even_only
# origin: languages/python/tests/python/test_filter_map_builtins.rs

list(filter(lambda x: x % 2 == 0, range(6)))
