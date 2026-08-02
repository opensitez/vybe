# vybe-test: python/python_threading_events/test_semaphore_acquire_release
# origin: languages/python/tests/python/test_python_threading_events.rs

import threading

sem = threading.Semaphore(2)
acquired = []

def task(n):
    if sem.acquire(timeout=1):
        acquired.append(n)
        sem.release()

threads = [threading.Thread(target=task, args=(i,)) for i in range(4)]
for t in threads:
    t.start()
for t in threads:
    t.join()

print(len(acquired) == 4)
