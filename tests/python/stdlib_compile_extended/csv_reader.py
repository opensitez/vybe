# vybe-test: python/stdlib_compile_extended/csv_reader
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# vybe-test-mode: compile

import csv
import io
csv.reader(io.StringIO('a,b\n1,2'))
