# vybe-test: python/list_methods_extended/list_sort_stable_equal_keys
# origin: languages/python/tests/python/test_list_methods_extended.rs

a = [(1, 'b'), (1, 'a'), (2, 'c')]
a.sort(key=lambda t: t[0])
print([x[1] for x in a])
