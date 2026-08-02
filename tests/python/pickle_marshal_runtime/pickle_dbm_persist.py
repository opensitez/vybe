# vybe-test: python/pickle_marshal_runtime/pickle_dbm_persist
# origin: languages/python/tests/python/test_pickle_marshal_runtime.rs
# vybe-test-mode: compile

import pickle
import io
pickle.Pickler(io.BytesIO())
