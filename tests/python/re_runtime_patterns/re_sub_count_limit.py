# vybe-test: python/re_runtime_patterns/re_sub_count_limit
# origin: languages/python/tests/python/test_re_runtime_patterns.rs

import re; re.sub('a', 'x', 'aaa', count=2)
