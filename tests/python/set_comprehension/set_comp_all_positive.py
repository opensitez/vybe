# vybe-test: python/set_comprehension/set_comp_all_positive
# origin: languages/python/tests/python/test_set_comprehension.rs

s = {x for x in range(1, 4)}
print(all(v > 0 for v in s))
