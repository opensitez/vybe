# vybe-test: python/lambda_core/lambda_sorted_key_second_element
# origin: languages/python/tests/python/test_lambda_core.rs

sorted([(2, 'b'), (1, 'a')], key=lambda t: t[1])
