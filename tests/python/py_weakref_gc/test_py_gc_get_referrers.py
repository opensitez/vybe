# vybe-test: python/py_weakref_gc/test_py_gc_get_referrers
# origin: languages/python/tests/python/test_py_weakref_gc.rs

import gc

class Tracked:
    pass

obj = Tracked()
container = [obj]

refs = gc.get_referrers(obj)
print(any(r is container for r in refs))
