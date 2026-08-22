# vybe-test: python/function_signatures_spec/variadic_kwonly_compile
# origin: languages/python/tests/python/test_function_signatures_spec.rs

def f(*args, sep=':', end='!'):
    return sep
