# vybe-test: python/py_heapq_bisect_array/test_py_heapq_nlargest_with_key
# origin: languages/python/tests/python/test_py_heapq_bisect_array.rs

import heapq

students = [
    {"name": "Alice", "gpa": 3.8},
    {"name": "Bob", "gpa": 3.5},
    {"name": "Charlie", "gpa": 3.9},
    {"name": "Dave", "gpa": 3.2},
]

top2 = heapq.nlargest(2, students, key=lambda s: s["gpa"])
print([s["name"] for s in top2])
