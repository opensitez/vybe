# vybe-test: python/new_features/os_rename
# origin: languages/python/tests/python/test_new_features.rs
import os as _os, tempfile as _tf
_os.chdir(_tf.mkdtemp())
open('old.txt', 'w').close()

import os
os.rename('old.txt', 'new.txt')
