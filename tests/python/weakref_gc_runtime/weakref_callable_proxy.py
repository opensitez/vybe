# vybe-test: python/weakref_gc_runtime/weakref_callable_proxy
# origin: languages/python/tests/python/test_weakref_gc_runtime.rs

import weakref
class C:
 def __call__(self):
  return 1
# The referent must be kept ALIVE: `weakref.proxy(C())` takes a TEMPORARY
# with no strong reference, so it is collected immediately and every use
# raises "weakly-referenced object no longer exists".
_alive = C()
p = weakref.proxy(_alive)
p()
