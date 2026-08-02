# vybe-test: python/lambda_core/lambda_in_list_comp
# origin: languages/python/tests/python/test_lambda_core.rs

[(lambda x: x + 1)(i) for i in range(3)]
