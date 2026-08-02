# vybe-test: python/io_runtime/open_with_statement
# origin: languages/python/tests/python/test_io_runtime.rs
# vybe-test-mode: compile

with open(__file__) as f:
 f.readline()
