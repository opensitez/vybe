# vybe-test: python/python_queue_thread_safe_datastructures/test_queue_clear_via_drain
# origin: languages/python/tests/python/test_python_queue_thread_safe_datastructures.rs

import queue
q = queue.Queue()
for i in range(5):
    q.put(i)
drained = []
while not q.empty():
    drained.append(q.get())
print(drained)
