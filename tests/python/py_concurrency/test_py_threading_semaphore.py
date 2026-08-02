# vybe-test: python/py_concurrency/test_py_threading_semaphore
# origin: languages/python/tests/python/test_py_concurrency.rs

import threading

sem = threading.Semaphore(2)
active = [0]
max_seen = [0]
lock = threading.Lock()

def task():
    with sem:
        with lock:
            active[0] += 1
            max_seen[0] = max(max_seen[0], active[0])
        import time; time.sleep(0.01)
        with lock:
            active[0] -= 1

threads = [threading.Thread(target=task) for _ in range(6)]
for t in threads:
    t.start()
for t in threads:
    t.join()

print(max_seen[0] <= 2)
