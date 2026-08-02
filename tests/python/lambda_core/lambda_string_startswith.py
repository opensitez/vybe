# vybe-test: python/lambda_core/lambda_string_startswith
# origin: languages/python/tests/python/test_lambda_core.rs

list(filter(lambda s: s.startswith('a'), ['ab', 'ba']))
