# vybe-test: python/py_concurrency/test_py_concurrent_futures_as_completed
# origin: languages/python/tests/python/test_py_concurrency.rs

from concurrent.futures import ThreadPoolExecutor, as_completed
import time

def slow(n):
    time.sleep(n * 0.01)
    return n

with ThreadPoolExecutor(max_workers=4) as ex:
    futures = {ex.submit(slow, i): i for i in [3, 1, 2]}
    order = []
    for f in as_completed(futures):
        order.append(f.result())

print(sorted(order))
