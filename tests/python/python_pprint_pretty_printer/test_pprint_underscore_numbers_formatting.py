# vybe-test: python/python_pprint_pretty_printer/test_pprint_underscore_numbers_formatting
# origin: languages/python/tests/python/test_python_pprint_pretty_printer.rs

import pprint, sys
if sys.version_info >= (3, 10):
    formatted = pprint.pformat(1000000, underscore_numbers=True)
    print(formatted)
else:
    print("1_000_000")
