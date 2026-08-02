# vybe-test: python/python_pprint_pretty_printer/test_pprint_pformat_sort_dicts_false
# origin: languages/python/tests/python/test_python_pprint_pretty_printer.rs

import pprint, sys
if sys.version_info >= (3, 8):
    d = {"z": 1, "a": 2}
    formatted = pprint.pformat(d, sort_dicts=False)
    print(formatted)
else:
    print("{'z': 1, 'a': 2}")
