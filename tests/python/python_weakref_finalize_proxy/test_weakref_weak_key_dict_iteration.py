# vybe-test: python/python_weakref_finalize_proxy/test_weakref_weak_key_dict_iteration
# origin: languages/python/tests/python/test_python_weakref_finalize_proxy.rs

import weakref

class Key:
    def __init__(self, k): self.k = k

d = weakref.WeakKeyDictionary()
k1 = Key("a")
k2 = Key("b")
d[k1] = 1
d[k2] = 2
keys = [k.k for k in d.keys()]
print(sorted(keys))
