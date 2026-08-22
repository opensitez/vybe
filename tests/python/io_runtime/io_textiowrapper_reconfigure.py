# vybe-test: python/io_runtime/io_textiowrapper_reconfigure
# origin: languages/python/tests/python/test_io_runtime.rs

import io
s = io.StringIO()
hasattr(s, 'reconfigure')
