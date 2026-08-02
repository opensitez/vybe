# vybe-test: python/re_extended_patterns/re_template_flag
# origin: languages/python/tests/python/test_re_extended_patterns.rs
# vybe-test-mode: compile

import re
re.findall('(a)', 'aba', re.T)
