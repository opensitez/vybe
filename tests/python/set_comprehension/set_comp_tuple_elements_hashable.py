# vybe-test: python/set_comprehension/set_comp_tuple_elements_hashable
# origin: languages/python/tests/python/test_set_comprehension.rs

sorted({(i, i + 1) for i in range(2)})
