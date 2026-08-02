# vybe-test: python/python_difflib_sequencematcher_diff/test_difflib_opcodes
# origin: languages/python/tests/python/test_python_difflib_sequencematcher_diff.rs

import difflib
sm = difflib.SequenceMatcher(None, "abc", "axc")
opcodes = sm.get_opcodes()
tag_names = [op[0] for op in opcodes]
print("equal" in tag_names)
print("replace" in tag_names)
