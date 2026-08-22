# vybe-test: python/weakref_gc_runtime/tracemalloc_compare_snapshots
# origin: languages/python/tests/python/test_weakref_gc_runtime.rs

import tracemalloc
tracemalloc.start()
s1 = tracemalloc.take_snapshot()
s2 = tracemalloc.take_snapshot()
tracemalloc.stop()
