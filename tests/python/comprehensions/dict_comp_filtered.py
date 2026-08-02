# vybe-test: python/comprehensions/dict_comp_filtered
# origin: languages/python/tests/python/test_comprehensions.rs
# vybe-test-mode: compile

d = {k: v for k, v in items.items() if v > 0}
