# vybe-test: python/pickle_marshal_runtime/marshal_empty_bytes_raises
# origin: languages/python/tests/python/test_pickle_marshal_runtime.rs

import marshal
try:
 marshal.loads(b'')
 print('ok')
except (ValueError, EOFError):
 print('err')
