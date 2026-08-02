# vybe-test: python/py_io_streams_buffers_bytes/test_py_io_line_iteration_readlines
# origin: languages/python/tests/python/test_py_io_streams_buffers_bytes.rs

import io

data = "line1\nline2\nline3"
buf = io.StringIO(data)
lines = [line.strip() for line in buf]
print(lines)
