# vybe-test: python/list_methods_extended/list_bisect_insert_at_start
# origin: languages/python/tests/python/test_list_methods_extended.rs

a = [2, 4, 6]
x = 0
i = 0
while i < len(a) and a[i] < x:
    i += 1
a.insert(i, x)
print(a)
