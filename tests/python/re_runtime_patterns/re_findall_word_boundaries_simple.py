# vybe-test: python/re_runtime_patterns/re_findall_word_boundaries_simple
# origin: languages/python/tests/python/test_re_runtime_patterns.rs

import re; re.findall(r'\b\w+\b', 'hi there')
