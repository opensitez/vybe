# vybe-test: python/re_runtime_patterns/re_sub_numeric_backref_in_pattern
# origin: languages/python/tests/python/test_re_runtime_patterns.rs

import re; re.sub(r'(a)(\d)', r'\2\1', 'a1')
