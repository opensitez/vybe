# vybe-test: python/io_runtime/open_binary_mode
# origin: languages/python/tests/python/test_io_runtime.rs
# vybe-test-mode: compile

with open(__file__, 'rb') as f:
 f.read(1)
