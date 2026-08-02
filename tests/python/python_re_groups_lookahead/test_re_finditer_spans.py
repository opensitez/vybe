# vybe-test: python/python_re_groups_lookahead/test_re_finditer_spans
# origin: languages/python/tests/python/test_python_re_groups_lookahead.rs

import re
spans = [(m.start(), m.end()) for m in re.finditer(r'\d+', 'abc123def456')]
print(spans)
