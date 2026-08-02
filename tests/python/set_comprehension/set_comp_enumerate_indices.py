# vybe-test: python/set_comprehension/set_comp_enumerate_indices
# origin: languages/python/tests/python/test_set_comprehension.rs

sorted({i for i, _ in enumerate(['a', 'b', 'a'])})
