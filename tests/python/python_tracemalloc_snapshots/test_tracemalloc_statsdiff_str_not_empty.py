# vybe-test: python/python_tracemalloc_snapshots/test_tracemalloc_statsdiff_str_not_empty
# origin: languages/python/tests/python/test_python_tracemalloc_snapshots.rs

import tracemalloc
tracemalloc.start()
snap1 = tracemalloc.take_snapshot()
x = [0] * 10000
snap2 = tracemalloc.take_snapshot()
diff = snap2.compare_to(snap1, "lineno")
if diff:
    print(len(str(diff[0])) > 0)
else:
    print(True)
tracemalloc.stop()
