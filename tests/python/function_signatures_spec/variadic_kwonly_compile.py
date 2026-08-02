# vybe-test: python/function_signatures_spec/variadic_kwonly_compile
# origin: languages/python/tests/python/test_function_signatures_spec.rs
# vybe-test-mode: compile

def f(*args, sep=':', end='!'):
    return sep
