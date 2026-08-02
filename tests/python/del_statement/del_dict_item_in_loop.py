# vybe-test: python/del_statement/del_dict_item_in_loop
# origin: languages/python/tests/python/test_del_statement.rs

d = {'a': 1, 'b': 2}
for k in list(d):
 if k == 'a':
  del d[k]
print(d)
