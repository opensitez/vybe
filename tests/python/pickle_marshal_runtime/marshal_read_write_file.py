# vybe-test: python/pickle_marshal_runtime/marshal_read_write_file
# origin: languages/python/tests/python/test_pickle_marshal_runtime.rs

import marshal
import tempfile
f = tempfile.NamedTemporaryFile()
marshal.dump(1, f)
