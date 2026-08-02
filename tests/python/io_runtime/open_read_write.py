# vybe-test: python/io_runtime/open_read_write
# origin: languages/python/tests/python/test_io_runtime.rs
# vybe-test-mode: compile

f = open(__file__)
data = f.read(10)
f.close()
