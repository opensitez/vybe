# vybe-test: python/py_string_regular_expressions_advanced/test_py_re_finditer_match_span_tuples
# origin: languages/python/tests/python/test_py_string_regular_expressions_advanced.rs

import re

text = "cat, bat, rat, mat"
matches = list(re.finditer(r"\b\w+at\b", text))
print([m.group() for m in matches])
print([m.span() for m in matches])
