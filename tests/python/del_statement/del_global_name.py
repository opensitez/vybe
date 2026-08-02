# vybe-test: python/del_statement/del_global_name
# origin: languages/python/tests/python/test_del_statement.rs

g = 1
def f():
 global g
 del g
f()
try:
 print(g)
except NameError:
 print('gone')
