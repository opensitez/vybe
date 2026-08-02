# vybe-test: python/list_comprehension/list_comp_with_condition_on_transform
# origin: languages/python/tests/python/test_list_comprehension.rs

[x * 2 for x in range(5) if x % 2 == 1]
