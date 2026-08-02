# vybe-test: python/list_comprehension/list_comp_enumerate_style
# origin: languages/python/tests/python/test_list_comprehension.rs

print([i for i, v in enumerate(['x', 'y']) if v == 'y'][0])
