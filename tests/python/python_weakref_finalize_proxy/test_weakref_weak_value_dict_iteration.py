# vybe-test: python/python_weakref_finalize_proxy/test_weakref_weak_value_dict_iteration
# origin: languages/python/tests/python/test_python_weakref_finalize_proxy.rs

import weakref

class Val:
    def __init__(self, v): self.v = v

d = weakref.WeakValueDictionary()
v1 = Val(1)
v2 = Val(2)
d["a"] = v1
d["b"] = v2
vals = [v.v for v in d.values()]
print(sorted(vals))
