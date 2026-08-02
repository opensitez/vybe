# vybe-test: python/weakref_gc_runtime/weakref_callable_proxy
# origin: languages/python/tests/python/test_weakref_gc_runtime.rs
# vybe-test-mode: compile

import weakref
class C:
 def __call__(self):
  return 1
p = weakref.proxy(C())
p()
