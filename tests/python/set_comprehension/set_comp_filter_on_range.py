# vybe-test: python/set_comprehension/set_comp_filter_on_range
# origin: languages/python/tests/python/test_set_comprehension.rs

sorted({x for x in filter(lambda n: n % 2 == 1, range(6))})
