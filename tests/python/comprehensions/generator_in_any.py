# vybe-test: python/comprehensions/generator_in_any
# origin: languages/python/tests/python/test_comprehensions.rs
data = [1, 2]

any_big = any(x > 100 for x in data)
