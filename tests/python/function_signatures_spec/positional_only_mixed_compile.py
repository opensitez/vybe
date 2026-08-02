# vybe-test: python/function_signatures_spec/positional_only_mixed_compile
# origin: languages/python/tests/python/test_function_signatures_spec.rs
# vybe-test-mode: compile

def add(a, b, /, c):
    return a + b + c
