# vybe-test: python/functions/dict_comp
# origin: languages/python/tests/python/test_functions.rs
# vybe-test-mode: compile

d = {k: k * 2 for k in range(5)}
