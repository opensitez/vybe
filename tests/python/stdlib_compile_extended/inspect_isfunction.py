# vybe-test: python/stdlib_compile_extended/inspect_isfunction
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# vybe-test-mode: compile

import inspect
def f(): pass
inspect.isfunction(f)
