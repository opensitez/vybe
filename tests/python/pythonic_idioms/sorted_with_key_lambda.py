# vybe-test: python/pythonic_idioms/sorted_with_key_lambda
# origin: languages/python/tests/python/test_pythonic_idioms.rs

sorted(['bb', 'a', 'ccc'], key=lambda s: len(s))
