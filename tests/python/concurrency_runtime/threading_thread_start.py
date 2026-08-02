# vybe-test: python/concurrency_runtime/threading_thread_start
# origin: languages/python/tests/python/test_concurrency_runtime.rs
# vybe-test-mode: compile

import threading
t = threading.Thread(target=lambda: None)
t.start()
t.join()
