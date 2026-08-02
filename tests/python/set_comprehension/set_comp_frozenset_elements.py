# vybe-test: python/set_comprehension/set_comp_frozenset_elements
# origin: languages/python/tests/python/test_set_comprehension.rs

sorted({frozenset({i}) for i in range(2)})
