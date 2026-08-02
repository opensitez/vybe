# vybe-test: python/data_structures/dict_unpacking
# origin: languages/python/tests/python/test_data_structures.rs
# vybe-test-mode: compile

a = {'x': 1}
b = {'y': 2}
c = {**a, **b}
