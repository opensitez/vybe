# vybe-test: python/set_comprehension/set_comp_from_list_removes_dupes
# origin: languages/python/tests/python/test_set_comprehension.rs

sorted({x for x in [1, 1, 2, 3, 2]})
