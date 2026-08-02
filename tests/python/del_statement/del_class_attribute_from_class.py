# vybe-test: python/del_statement/del_class_attribute_from_class
# origin: languages/python/tests/python/test_del_statement.rs

class C:
 x = 1
del C.x
try:
 print(C.x)
except AttributeError:
 print('attr')
