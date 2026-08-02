# vybe-test: python/python_io_stringio_bytesio/test_stringio_getvalue
# origin: languages/python/tests/python/test_python_io_stringio_bytesio.rs

import io
buf = io.StringIO()
buf.write("abc\ndef\n")
print(buf.getvalue())
