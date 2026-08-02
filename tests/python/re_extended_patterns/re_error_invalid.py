# vybe-test: python/re_extended_patterns/re_error_invalid
# origin: languages/python/tests/python/test_re_extended_patterns.rs

import re
try:
 re.compile('(')
 print('ok')
except re.error:
 print('bad')
