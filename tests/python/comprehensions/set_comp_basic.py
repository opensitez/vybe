# vybe-test: python/comprehensions/set_comp_basic
# origin: languages/python/tests/python/test_comprehensions.rs
# vybe-test-mode: compile

s = {x * x for x in range(5)}
