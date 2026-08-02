# vybe-test: python/struct_copy_encoding/pickle_dumps_loads
# origin: languages/python/tests/python/test_struct_copy_encoding.rs
# vybe-test-mode: compile

import pickle
pickle.loads(pickle.dumps([1, 2]))
