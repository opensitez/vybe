# vybe-test: python/try_except_core/try_except_base_exception_not_caught_by_exception
# origin: languages/python/tests/python/test_try_except_core.rs

try:
 raise KeyboardInterrupt
except Exception:
 print('exc')
except BaseException:
 print('base')
