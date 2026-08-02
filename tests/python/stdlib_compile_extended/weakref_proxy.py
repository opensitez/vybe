# vybe-test: python/stdlib_compile_extended/weakref_proxy
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# vybe-test-mode: compile

import weakref
class C: pass
weakref.proxy(C())
