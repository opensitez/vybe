# vybe-test: python/python_tracemalloc_snapshots/test_tracemalloc_stat_size_and_count
# origin: languages/python/tests/python/test_python_tracemalloc_snapshots.rs

import tracemalloc
tracemalloc.start()
data = b"x" * 100000
snap = tracemalloc.take_snapshot()
stats = snap.statistics("lineno")
total_size = sum(s.size for s in stats)
total_count = sum(s.count for s in stats)
print(total_size > 0)
print(total_count > 0)
tracemalloc.stop()
