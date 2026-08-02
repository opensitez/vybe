# vybe-test: python/py_heapq_bisect_array/test_py_heapq_basic_heap_operations
# origin: languages/python/tests/python/test_py_heapq_bisect_array.rs

import heapq

heap = []
for val in [5, 1, 8, 3, 9, 2]:
    heapq.heappush(heap, val)

print(heap[0])   # smallest element always at index 0
results = []
while heap:
    results.append(heapq.heappop(heap))
print(results)
