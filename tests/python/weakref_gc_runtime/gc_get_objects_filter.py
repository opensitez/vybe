# vybe-test: python/weakref_gc_runtime/gc_get_objects_filter
# origin: languages/python/tests/python/test_weakref_gc_runtime.rs
# vybe-test-mode: compile

import gc
[o for o in gc.get_objects() if isinstance(o, list)][:1]
