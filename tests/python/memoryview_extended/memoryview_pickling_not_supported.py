# vybe-test: python/memoryview_extended/memoryview_pickling_not_supported
# origin: languages/python/tests/python/test_memoryview_extended.rs

import pickle
try:
 pickle.dumps(memoryview(b'a'))
 print('ok')
except TypeError:
 print('err')
