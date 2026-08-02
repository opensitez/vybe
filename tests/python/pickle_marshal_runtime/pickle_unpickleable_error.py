# vybe-test: python/pickle_marshal_runtime/pickle_unpickleable_error
# origin: languages/python/tests/python/test_pickle_marshal_runtime.rs

import pickle
class C:
 pass
try:
 pickle.dumps(C())
 print('ok')
except (pickle.PicklingError, TypeError):
 print('err')
