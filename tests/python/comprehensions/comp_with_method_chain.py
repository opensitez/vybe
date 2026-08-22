# vybe-test: python/comprehensions/comp_with_method_chain
# origin: languages/python/tests/python/test_comprehensions.rs
# The base/name this fixture uses was never defined — supplied so it RUNS.
lines = [' A ', ' B ']


result = [word.strip().lower() for word in lines]
