# vybe-test: python/py_concurrency/test_py_threading_lock_mutual_exclusion
# origin: languages/python/tests/python/test_py_concurrency.rs

import threading

counter = [0]
lock = threading.Lock()

def increment():
    for _ in range(100):
        with lock:
            counter[0] += 1

threads = [threading.Thread(target=increment) for _ in range(5)]
for t in threads:
    t.start()
for t in threads:
    t.join()

print(counter[0])
