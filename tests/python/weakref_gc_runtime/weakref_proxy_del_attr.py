# vybe-test: python/weakref_gc_runtime/weakref_proxy_del_attr
# origin: languages/python/tests/python/test_weakref_gc_runtime.rs
# vybe-test-mode: compile

import weakref
class C:
 x = 1
p = weakref.proxy(C())
del p.x
