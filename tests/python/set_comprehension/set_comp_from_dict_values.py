# vybe-test: python/set_comprehension/set_comp_from_dict_values
# origin: languages/python/tests/python/test_set_comprehension.rs

sorted({v for v in {'a': 1, 'b': 2}.values()})
