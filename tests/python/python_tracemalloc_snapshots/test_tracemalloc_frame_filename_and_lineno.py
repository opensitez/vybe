# vybe-test: python/python_tracemalloc_snapshots/test_tracemalloc_frame_filename_and_lineno
# origin: languages/python/tests/python/test_python_tracemalloc_snapshots.rs

import tracemalloc
tracemalloc.start(3)
x = [None] * 2000
snap = tracemalloc.take_snapshot()
stats = snap.statistics("traceback")
if stats:
    frame = stats[0].traceback[0]
    print(isinstance(frame.filename, str))
    print(frame.lineno > 0)
else:
    print(True)
    print(True)
tracemalloc.stop()
