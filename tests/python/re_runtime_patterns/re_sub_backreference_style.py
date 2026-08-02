# vybe-test: python/re_runtime_patterns/re_sub_backreference_style
# origin: languages/python/tests/python/test_re_runtime_patterns.rs

import re; re.sub(r'(a)(b)', r'\2\1', 'ab')
