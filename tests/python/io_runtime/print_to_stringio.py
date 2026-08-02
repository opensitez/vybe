# vybe-test: python/io_runtime/print_to_stringio
# origin: languages/python/tests/python/test_io_runtime.rs

import io
import sys
buf = io.StringIO()
print('hi', file=buf)
print(buf.getvalue().strip())
