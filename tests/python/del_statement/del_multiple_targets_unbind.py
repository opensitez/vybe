# vybe-test: python/del_statement/del_multiple_targets_unbind
# origin: languages/python/tests/python/test_del_statement.rs

a = b = 1
del a, b
try:
 print(a)
except NameError:
 print('ok')
