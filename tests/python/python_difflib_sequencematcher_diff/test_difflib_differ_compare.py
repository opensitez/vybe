# vybe-test: python/python_difflib_sequencematcher_diff/test_difflib_differ_compare
# origin: languages/python/tests/python/test_python_difflib_sequencematcher_diff.rs

import difflib
d = difflib.Differ()
res = list(d.compare(["a", "b"], ["a", "c"]))
print([line.strip() for line in res])
