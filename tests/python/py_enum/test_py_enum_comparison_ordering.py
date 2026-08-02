# vybe-test: python/py_enum/test_py_enum_comparison_ordering
# origin: languages/python/tests/python/test_py_enum.rs

from enum import IntEnum

class Priority(IntEnum):
    LOW = 1
    MEDIUM = 2
    HIGH = 3

items = [("task_a", Priority.HIGH), ("task_b", Priority.LOW), ("task_c", Priority.MEDIUM)]
sorted_items = sorted(items, key=lambda x: x[1])
print([name for name, _ in sorted_items])
