# vybe-test: python/del_statement/del_from_list_while_iterating_copy
# origin: languages/python/tests/python/test_del_statement.rs

a = [1, 2, 3]
for x in list(a):
 if x == 2:
  a.remove(x)
print(a)
