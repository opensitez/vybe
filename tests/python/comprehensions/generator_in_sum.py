# vybe-test: python/comprehensions/generator_in_sum
# origin: languages/python/tests/python/test_comprehensions.rs

total = sum(x * x for x in range(10))
