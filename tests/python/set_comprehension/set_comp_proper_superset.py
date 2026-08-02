# vybe-test: python/set_comprehension/set_comp_proper_superset
# origin: languages/python/tests/python/test_set_comprehension.rs

a = {x for x in range(3)}
b = {x for x in range(2)}
print(a > b)
