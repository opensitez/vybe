# vybe-test: python/lambda_core/lambda_filter_truthy_strings
# origin: languages/python/tests/python/test_lambda_core.rs

list(filter(lambda s: bool(s), ['', 'a', '']))
