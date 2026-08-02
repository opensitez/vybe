# vybe-test: python/for_while_extended/for_list_mutate
# origin: languages/python/tests/python/test_for_while_extended.rs

a = [1, 2, 3]
for i, v in enumerate(a):
 a[i] = v * 2
print(a)
