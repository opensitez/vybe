# vybe-test: python/set_comprehension/set_comp_filtered_evens
# origin: languages/python/tests/python/test_set_comprehension.rs

sorted({x for x in range(6) if x % 2 == 0})
