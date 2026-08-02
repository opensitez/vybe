# vybe-test: python/comprehensions/generator_in_join
# origin: languages/python/tests/python/test_comprehensions.rs
# vybe-test-mode: compile

joined = ','.join(str(x) for x in [1,2,3])
