# vybe-test: python/csv_markup_runtime/csv_register_dialect
# origin: languages/python/tests/python/test_csv_markup_runtime.rs
# vybe-test-mode: compile

import csv
csv.register_dialect('x', delimiter=';')
