# vybe-test: python/py_concurrency/test_py_threading_queue_producer_consumer
# origin: languages/python/tests/python/test_py_concurrency.rs

import threading
from queue import Queue

q = Queue()
results = []

def producer():
    for i in range(5):
        q.put(i)
    q.put(None)  # sentinel

def consumer():
    while True:
        item = q.get()
        if item is None:
            break
        results.append(item)

p = threading.Thread(target=producer)
c = threading.Thread(target=consumer)
p.start()
c.start()
p.join()
c.join()
print(results)
