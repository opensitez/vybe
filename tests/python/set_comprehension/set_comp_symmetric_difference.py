# vybe-test: python/set_comprehension/set_comp_symmetric_difference
# origin: languages/python/tests/python/test_set_comprehension.rs

sorted({x for x in range(3)} ^ {1, 2, 4})
