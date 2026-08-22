# vybe-test: python/function_signatures_spec/positional_only_basic_compile
# origin: languages/python/tests/python/test_function_signatures_spec.rs

def add(a, b, /):
    return a + b
