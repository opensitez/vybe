# vybe-test: python/exceptions_extended/except_else_skipped_on_error
# origin: languages/python/tests/python/test_exceptions_extended.rs

try:
 raise ValueError()
except ValueError:
 print('caught')
else:
 print('else')
