# vybe-test: python/python_tracemalloc_snapshots/test_tracemalloc_statsdiff_size_delta_positive
# origin: languages/python/tests/python/test_python_tracemalloc_snapshots.rs

import tracemalloc
tracemalloc.start()
snap1 = tracemalloc.take_snapshot()
alloc = bytearray(500000)
snap2 = tracemalloc.take_snapshot()
diff = snap2.compare_to(snap1, "lineno")
# At least one stat shows positive size delta
print(any(d.size_diff > 0 for d in diff))
tracemalloc.stop()
