# vybe-test: python/python_pprint_pretty_printer/test_pprint_pformat_width_wrapping
# origin: languages/python/tests/python/test_python_pprint_pretty_printer.rs

import pprint
items = ["item_" + str(i) for i in range(10)]
formatted = pprint.pformat(items, width=20)
print("\n" in formatted)
