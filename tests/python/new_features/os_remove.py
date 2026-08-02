# vybe-test: python/new_features/os_remove
# origin: languages/python/tests/python/test_new_features.rs
# vybe-test-mode: compile

import os
os.remove('/tmp/test.txt')
