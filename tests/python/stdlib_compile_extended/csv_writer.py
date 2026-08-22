# vybe-test: python/stdlib_compile_extended/csv_writer
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs

import csv
import io
csv.writer(io.StringIO())
