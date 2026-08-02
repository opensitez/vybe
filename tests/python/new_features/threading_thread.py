# vybe-test: python/new_features/threading_thread
# origin: languages/python/tests/python/test_new_features.rs
# vybe-test-mode: compile

import threading
def worker():
    print('hello from thread')
t = threading.Thread(worker)
