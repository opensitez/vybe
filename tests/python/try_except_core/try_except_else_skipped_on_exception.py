# vybe-test: python/try_except_core/try_except_else_skipped_on_exception
# origin: languages/python/tests/python/test_try_except_core.rs

try:
 1/0
except ZeroDivisionError:
 print('ex')
else:
 print('else')
