# vybe-test: python/del_statement/del_bytearray_item
# origin: languages/python/tests/python/test_del_statement.rs

try:
 ba = bytearray(b'ab')
 del ba[0]
 print(ba)
except:
 print('err')
