# vybe-test: python/try_except_core/try_except_specific_before_general
# origin: languages/python/tests/python/test_try_except_core.rs

try:
 int('x')
except ValueError:
 print('specific')
except Exception:
 print('general')
