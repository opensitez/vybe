# vybe-test: python/stdlib_compile_extended/weakref_ref
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs

import weakref
class C: pass
weakref.ref(C())
