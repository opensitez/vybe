# vybe-test: python/python_pprint_pretty_printer/test_pprint_pp_alias_helper
# origin: languages/python/tests/python/test_python_pprint_pretty_printer.rs

import pprint, io, sys
if sys.version_info >= (3, 8):
    buf = io.StringIO()
    pprint.pp({"b": 1, "a": 2}, stream=buf)
    print(buf.getvalue().strip())
else:
    print("{'a': 2, 'b': 1}")
