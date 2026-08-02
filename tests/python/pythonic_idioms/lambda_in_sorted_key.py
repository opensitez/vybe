# vybe-test: python/pythonic_idioms/lambda_in_sorted_key
# origin: languages/python/tests/python/test_pythonic_idioms.rs

sorted([(2, 'b'), (1, 'a')], key=lambda t: t[1])
