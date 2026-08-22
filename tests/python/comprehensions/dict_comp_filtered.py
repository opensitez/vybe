# vybe-test: python/comprehensions/dict_comp_filtered
# origin: languages/python/tests/python/test_comprehensions.rs
items = {'key': 1, 'a': 1}

d = {k: v for k, v in items.items() if v > 0}
