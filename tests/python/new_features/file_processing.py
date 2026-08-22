# vybe-test: python/new_features/file_processing
# origin: languages/python/tests/python/test_new_features.rs

import os
files = os.listdir('.')
for f in files:
    if os.path.isfile(f):
        size = os.path.getsize(f)
        print(f, size)
