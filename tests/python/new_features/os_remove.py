# vybe-test: python/new_features/os_remove
# origin: languages/python/tests/python/test_new_features.rs
import os as _os, tempfile as _tf
_os.chdir(_tf.mkdtemp())
open('/tmp/test.txt', 'w').close()

import os
os.remove('/tmp/test.txt')
