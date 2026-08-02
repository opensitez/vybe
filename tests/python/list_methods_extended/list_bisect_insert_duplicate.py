# vybe-test: python/list_methods_extended/list_bisect_insert_duplicate
# origin: languages/python/tests/python/test_list_methods_extended.rs

a = [1, 2, 2, 3]
x = 2
i = 0
while i < len(a) and a[i] <= x:
    i += 1
a.insert(i, x)
print(a)
