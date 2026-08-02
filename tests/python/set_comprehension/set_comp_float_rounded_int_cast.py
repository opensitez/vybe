# vybe-test: python/set_comprehension/set_comp_float_rounded_int_cast
# origin: languages/python/tests/python/test_set_comprehension.rs

sorted({int(x) for x in [1.1, 1.9, 2.1]})
