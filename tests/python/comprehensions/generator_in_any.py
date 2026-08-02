# vybe-test: python/comprehensions/generator_in_any
# origin: languages/python/tests/python/test_comprehensions.rs
# vybe-test-mode: compile

any_big = any(x > 100 for x in data)
