# vybe-test: python/python_queue_thread_safe_datastructures/test_priority_queue_custom_comparable
# origin: languages/python/tests/python/test_python_queue_thread_safe_datastructures.rs

import queue
q = queue.PriorityQueue()
q.put((10, "task_b"))
q.put((5, "task_a"))
while not q.empty():
    priority, task = q.get()
    print(f"{task}:{priority}")
