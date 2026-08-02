# vybe-test: python/re_extended_patterns/re_unicode_flag
# origin: languages/python/tests/python/test_re_extended_patterns.rs
# vybe-test-mode: compile

import re
re.compile(r'\w', re.U)
