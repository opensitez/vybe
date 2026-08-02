# vybe-test: python/new_features/re_sub
# origin: languages/python/tests/python/test_new_features.rs
# vybe-test-mode: compile

import re
s = re.sub(r'\d', 'X', 'a1b2')
