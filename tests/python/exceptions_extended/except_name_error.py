# vybe-test: python/exceptions_extended/except_name_error
# origin: languages/python/tests/python/test_exceptions_extended.rs

try:
 print(undefined_name_xyz)
except NameError:
 print('ne')
