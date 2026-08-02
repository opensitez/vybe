# vybe-test: python/stdlib_compile_extended/inspect_getmembers
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# vybe-test-mode: compile

import inspect
class C: pass
inspect.getmembers(C)
