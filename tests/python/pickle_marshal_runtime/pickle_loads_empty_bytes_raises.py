# vybe-test: python/pickle_marshal_runtime/pickle_loads_empty_bytes_raises
# origin: languages/python/tests/python/test_pickle_marshal_runtime.rs

import pickle
try:
 pickle.loads(b'')
 print('ok')
except pickle.UnpicklingError:
 print('err')
