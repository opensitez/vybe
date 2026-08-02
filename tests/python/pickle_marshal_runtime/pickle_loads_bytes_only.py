# vybe-test: python/pickle_marshal_runtime/pickle_loads_bytes_only
# origin: languages/python/tests/python/test_pickle_marshal_runtime.rs

import pickle
try:
 pickle.loads('not bytes')
 print('ok')
except (TypeError, pickle.UnpicklingError):
 print('err')
