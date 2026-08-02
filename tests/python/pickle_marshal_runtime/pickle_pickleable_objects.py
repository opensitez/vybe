# vybe-test: python/pickle_marshal_runtime/pickle_pickleable_objects
# origin: languages/python/tests/python/test_pickle_marshal_runtime.rs
# vybe-test-mode: compile

import pickle
class C:
 pass
try:
 pickle.dumps(C())
except:
 pass
