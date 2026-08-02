# vybe-test: python/lambda_core/lambda_sorted_key_length
# origin: languages/python/tests/python/test_lambda_core.rs

sorted(['bb', 'a', 'ccc'], key=lambda s: len(s))
