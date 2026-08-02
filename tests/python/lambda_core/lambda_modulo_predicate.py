# vybe-test: python/lambda_core/lambda_modulo_predicate
# origin: languages/python/tests/python/test_lambda_core.rs

list(filter(lambda x: x % 2 == 0, range(6)))
