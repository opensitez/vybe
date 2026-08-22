# vybe-test: python/function_signatures_spec/variadic_positional_only_compile
# origin: languages/python/tests/python/test_function_signatures_spec.rs

def f(a, /, *args):
    return args
