# vybe-test: python/py_concurrency/test_py_threading_basic_thread
# origin: languages/python/tests/python/test_py_concurrency.rs

import threading

results = []

def worker(n):
    results.append(n * n)

threads = [threading.Thread(target=worker, args=(i,)) for i in range(4)]
for t in threads:
    t.start()
for t in threads:
    t.join()

print(sorted(results))
