# vybe-test: python/comprehensions/dict_comp_from_list
# origin: languages/python/tests/python/test_comprehensions.rs
# vybe-test-mode: compile

d = {s: len(s) for s in ['hello', 'world']}
