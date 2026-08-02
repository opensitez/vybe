# vybe-test: python/list_comprehension/list_comp_negative_indices_via_enumerate
# origin: languages/python/tests/python/test_list_comprehension.rs

[i for i, v in enumerate([10, 20, 30]) if v == 30]
