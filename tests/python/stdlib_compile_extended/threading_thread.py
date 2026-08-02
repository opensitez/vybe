# vybe-test: python/stdlib_compile_extended/threading_thread
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# vybe-test-mode: compile

import threading
t = threading.Thread(target=lambda: None)
