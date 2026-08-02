# vybe-test: python/pickle_marshal_runtime/pickle_roundtrip_range_not_picklable
# origin: languages/python/tests/python/test_pickle_marshal_runtime.rs

import pickle
try:
 pickle.dumps(range(3))
 print('ok')
except (TypeError, pickle.PicklingError):
 print('err')
