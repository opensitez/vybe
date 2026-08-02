# vybe-test: python/py_threading_synchronization_primitives/test_py_threading_barrier_synchronization
# origin: languages/python/tests/python/test_py_threading_synchronization_primitives.rs

import threading

barrier = threading.Barrier(3)
log = []

def worker(idx):
    log.append(f"w{idx}_before")
    barrier.wait()
    log.append(f"w{idx}_after")

threads = [threading.Thread(target=worker, args=(i,)) for i in range(3)]
for t in threads: t.start()
for t in threads: t.join()

print(len([x for x in log if "after" in x]) == 3)
