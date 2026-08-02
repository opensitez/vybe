# vybe-test: python/py_concurrency/test_py_threading_barrier
# origin: languages/python/tests/python/test_py_concurrency.rs

import threading

log = []
barrier = threading.Barrier(3)

def task(name):
    log.append(f"{name}_before")
    barrier.wait()
    log.append(f"{name}_after")

threads = [threading.Thread(target=task, args=(f"T{i}",)) for i in range(3)]
for t in threads:
    t.start()
for t in threads:
    t.join()

befores = sorted(x for x in log if "before" in x)
afters = sorted(x for x in log if "after" in x)
print(befores)
print(afters)
