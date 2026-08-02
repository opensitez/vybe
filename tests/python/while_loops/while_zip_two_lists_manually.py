# vybe-test: python/while_loops/while_zip_two_lists_manually
# origin: languages/python/tests/python/test_while_loops.rs

a = [1, 2]
b = [3, 4]
i = 0
while i < len(a):
 print(a[i] + b[i])
 i += 1
