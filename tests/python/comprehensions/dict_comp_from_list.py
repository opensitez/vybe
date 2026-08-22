# vybe-test: python/comprehensions/dict_comp_from_list
# origin: languages/python/tests/python/test_comprehensions.rs

d = {s: len(s) for s in ['hello', 'world']}
