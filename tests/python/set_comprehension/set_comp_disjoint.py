# vybe-test: python/set_comprehension/set_comp_disjoint
# origin: languages/python/tests/python/test_set_comprehension.rs

a = {x for x in range(2)}
b = {x for x in range(2, 4)}
print(a.isdisjoint(b))
