# vybe-test: python/struct_copy_encoding/pickle_highest_protocol
# origin: languages/python/tests/python/test_struct_copy_encoding.rs

import pickle
pickle.dumps([], protocol=pickle.HIGHEST_PROTOCOL)
