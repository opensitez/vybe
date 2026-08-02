# vybe-test: python/set_comprehension/set_comp_intersection_with_range
# origin: languages/python/tests/python/test_set_comprehension.rs

sorted({x for x in range(5)} & {2, 3, 9})
