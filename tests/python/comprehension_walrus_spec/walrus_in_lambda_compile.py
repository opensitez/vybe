# vybe-test: python/comprehension_walrus_spec/walrus_in_lambda_compile
# origin: languages/python/tests/python/test_comprehension_walrus_spec.rs

func = lambda x: (y := x + 1)
