# vybe-test: python/set_comprehension/set_comp_nested_loops
# origin: languages/python/tests/python/test_set_comprehension.rs

sorted({a + b for a in [1, 2] for b in [10, 20]})
