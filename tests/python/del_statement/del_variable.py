# vybe-test: python/del_statement/del_variable
# origin: languages/python/tests/python/test_del_statement.rs

x = 1
del x
try:
 print(x)
except NameError:
 print('gone')
