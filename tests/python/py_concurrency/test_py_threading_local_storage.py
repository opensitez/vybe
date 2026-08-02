# vybe-test: python/py_concurrency/test_py_threading_local_storage
# origin: languages/python/tests/python/test_py_concurrency.rs

import threading

local = threading.local()
results = {}

def task(name, val):
    local.value = val
    import time; time.sleep(0.005)
    results[name] = local.value

threads = [
    threading.Thread(target=task, args=("A", 1)),
    threading.Thread(target=task, args=("B", 2)),
    threading.Thread(target=task, args=("C", 3)),
]
for t in threads:
    t.start()
for t in threads:
    t.join()

print(results["A"])
print(results["B"])
print(results["C"])
