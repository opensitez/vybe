# vybe-test: python/function_signatures_spec/variadic_positional_only_compile
# origin: languages/python/tests/python/test_function_signatures_spec.rs
# vybe-test-mode: compile

def f(a, /, *args):
    return args
