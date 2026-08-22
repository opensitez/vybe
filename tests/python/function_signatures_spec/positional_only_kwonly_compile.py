# vybe-test: python/function_signatures_spec/positional_only_kwonly_compile
# origin: languages/python/tests/python/test_function_signatures_spec.rs

def f(a, /, *, b):
    return a + b
