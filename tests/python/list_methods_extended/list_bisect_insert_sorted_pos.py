# vybe-test: python/list_methods_extended/list_bisect_insert_sorted_pos
# origin: languages/python/tests/python/test_list_methods_extended.rs

a = [1, 3, 5]
x = 4
i = 0
while i < len(a) and a[i] < x:
    i += 1
a.insert(i, x)
print(a)
