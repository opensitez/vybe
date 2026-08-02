# vybe-test: python/new_features/os_path_isfile
# origin: languages/python/tests/python/test_new_features.rs
# vybe-test-mode: compile

import os
x = os.path.isfile('/tmp/test.txt')
