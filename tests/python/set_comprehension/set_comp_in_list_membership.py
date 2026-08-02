# vybe-test: python/set_comprehension/set_comp_in_list_membership
# origin: languages/python/tests/python/test_set_comprehension.rs

s = {x for x in [1, 2]}
print(2 in s)
