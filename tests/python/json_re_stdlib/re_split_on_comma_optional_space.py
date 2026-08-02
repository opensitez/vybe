# vybe-test: python/json_re_stdlib/re_split_on_comma_optional_space
# origin: languages/python/tests/python/test_json_re_stdlib.rs

import re; re.split(r',\s*', 'a, b ,c')
