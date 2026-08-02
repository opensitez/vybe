# vybe-test: python/py_threading_synchronization_primitives/test_py_threading_lock_counter_protection
# origin: languages/python/tests/python/test_py_threading_synchronization_primitives.rs

import threading

counter = 0
lock = threading.Lock()

def worker():
    global counter
    for _ in range(100):
        with lock:
            counter += 1

threads = [threading.Thread(target=worker) for _ in range(5)]
for t in threads: t.start()
for t in threads: t.join()

print(counter)
