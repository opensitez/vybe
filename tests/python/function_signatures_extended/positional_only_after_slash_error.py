# vybe-test: python/function_signatures_extended/positional_only_after_slash_error
# origin: languages/python/tests/python/test_function_signatures_extended.rs
# vybe-test-mode: compile

def f(a, /, /, b): pass
