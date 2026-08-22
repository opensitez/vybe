# vybe-test: python/stdlib_compile_extended/weakref_finalize
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs

import weakref
class C: pass
weakref.finalize(C(), print)
