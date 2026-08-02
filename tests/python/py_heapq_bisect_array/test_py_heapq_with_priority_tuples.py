# vybe-test: python/py_heapq_bisect_array/test_py_heapq_with_priority_tuples
# origin: languages/python/tests/python/test_py_heapq_bisect_array.rs

import heapq

# Priority queue using (priority, item) tuples
pq = []
heapq.heappush(pq, (3, "low"))
heapq.heappush(pq, (1, "high"))
heapq.heappush(pq, (2, "medium"))

results = []
while pq:
    priority, item = heapq.heappop(pq)
    results.append(item)
print(results)
