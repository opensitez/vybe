# vybe-test: python/new_features/os_rename
# origin: languages/python/tests/python/test_new_features.rs
# vybe-test-mode: compile

import os
os.rename('old.txt', 'new.txt')
