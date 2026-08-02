# vybe-test: python/set_comprehension/set_comp_truthy_values
# origin: languages/python/tests/python/test_set_comprehension.rs

sorted({x for x in [0, 1, 2, 0] if x})
