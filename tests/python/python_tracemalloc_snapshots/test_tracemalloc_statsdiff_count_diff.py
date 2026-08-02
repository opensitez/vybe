# vybe-test: python/python_tracemalloc_snapshots/test_tracemalloc_statsdiff_count_diff
# origin: languages/python/tests/python/test_python_tracemalloc_snapshots.rs

import tracemalloc
tracemalloc.start()
snap1 = tracemalloc.take_snapshot()
new_objects = [object() for _ in range(1000)]
snap2 = tracemalloc.take_snapshot()
diff = snap2.compare_to(snap1, "lineno")
print(any(d.count_diff > 0 for d in diff))
tracemalloc.stop()
