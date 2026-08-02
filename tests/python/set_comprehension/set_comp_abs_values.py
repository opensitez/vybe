# vybe-test: python/set_comprehension/set_comp_abs_values
# origin: languages/python/tests/python/test_set_comprehension.rs

sorted({abs(x) for x in [-2, -1, 1]})
