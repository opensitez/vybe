# vybe-test: python/set_comprehension/set_comp_enumerate_values
# origin: languages/python/tests/python/test_set_comprehension.rs

sorted({v for _, v in enumerate(['x', 'y'])})
