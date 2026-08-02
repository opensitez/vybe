# vybe-test: python/stdlib_compile_extended/inspect_isclass
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# vybe-test-mode: compile

import inspect
class C: pass
inspect.isclass(C)
