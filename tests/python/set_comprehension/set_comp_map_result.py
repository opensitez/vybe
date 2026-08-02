# vybe-test: python/set_comprehension/set_comp_map_result
# origin: languages/python/tests/python/test_set_comprehension.rs

sorted({x * 2 for x in map(int, ['1', '2', '2'])})
