# vybe-test: python/try_except_core/try_except_else_runs_when_no_error
# origin: languages/python/tests/python/test_try_except_core.rs

try:
 x = 1
except:
 print('no')
else:
 print('else')
