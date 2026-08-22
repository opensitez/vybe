# vybe-test: python/weakref_gc_runtime/weakref_proxy_del_attr
# origin: languages/python/tests/python/test_weakref_gc_runtime.rs

import weakref
class C:
 x = 1
# The referent must be kept ALIVE: `weakref.proxy(C())` takes a TEMPORARY
# with no strong reference, so it is collected immediately and every use
# raises "weakly-referenced object no longer exists".
_alive = C()
p = weakref.proxy(_alive)
# `x` is a CLASS attribute; deleting it through the proxy targets the
# INSTANCE, which has none. Give the instance its own first.
_alive.x = 2
del p.x
