# vybe-test: python/io_runtime/open_write_mode
# origin: languages/python/tests/python/test_io_runtime.rs

import tempfile
import os
p = tempfile.mktemp()
f = open(p, 'w')
f.write('x')
f.close()
os.remove(p)
