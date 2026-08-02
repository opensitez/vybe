# vybe-test: python/set_comprehension/set_comp_negative_numbers
# origin: languages/python/tests/python/test_set_comprehension.rs

sorted({x for x in [-1, -2, 1] if x < 0})
