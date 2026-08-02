# vybe-test: python/comprehensions/dict_comp_basic
# origin: languages/python/tests/python/test_comprehensions.rs
# vybe-test-mode: compile

d = {k: k*2 for k in range(5)}
