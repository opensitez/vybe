# vybe-test: python/py_threading_synchronization_primitives/test_py_threading_semaphore_max_concurrency
# origin: languages/python/tests/python/test_py_threading_synchronization_primitives.rs

import threading

sem = threading.Semaphore(2)
active = 0
max_active = 0
lock = threading.Lock()

def worker():
    global active, max_active
    with sem:
        with lock:
            active += 1
            if active > max_active: max_active = active
        import time; time.sleep(0.005)
        with lock:
            active -= 1

threads = [threading.Thread(target=worker) for _ in range(4)]
for t in threads: t.start()
for t in threads: t.join()

print(max_active <= 2)
