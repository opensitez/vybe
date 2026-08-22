# vybe-test: python/stdlib_compile_extended/inspect_signature
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs

import inspect
def f(a, b=1): pass
inspect.signature(f)
