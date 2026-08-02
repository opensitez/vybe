# vybe-test: python/py_weakref_gc/test_py_gc_collect_and_get_stats
# origin: languages/python/tests/python/test_py_weakref_gc.rs

import gc

gc.collect()  # run collection
stats = gc.get_stats()
print(len(stats))  # 3 generations
print(all("collections" in s for s in stats))
print(gc.isenabled())
