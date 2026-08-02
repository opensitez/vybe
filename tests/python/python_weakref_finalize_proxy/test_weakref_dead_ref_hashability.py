# vybe-test: python/python_weakref_finalize_proxy/test_weakref_dead_ref_hashability
# origin: languages/python/tests/python/test_python_weakref_finalize_proxy.rs

import weakref

class Dummy: pass

d = Dummy()
r = weakref.ref(d)
h1 = hash(r)
del d
try:
    h2 = hash(r)
    print(h1 == h2)
except TypeError:
    print("TypeError")
