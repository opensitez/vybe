# vybe-test: python/python_tracemalloc_snapshots/test_tracemalloc_stat_traceback_has_frames
# origin: languages/python/tests/python/test_python_tracemalloc_snapshots.rs

import tracemalloc
tracemalloc.start(5)
x = [None] * 5000
snap = tracemalloc.take_snapshot()
stats = snap.statistics("traceback")
if stats:
    tb = stats[0].traceback
    print(len(tb) > 0)
else:
    print(True)
tracemalloc.stop()
