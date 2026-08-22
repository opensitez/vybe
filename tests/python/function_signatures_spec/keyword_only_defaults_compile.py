# vybe-test: python/function_signatures_spec/keyword_only_defaults_compile
# origin: languages/python/tests/python/test_function_signatures_spec.rs

def f(a, *, left=1, right=2):
    return a + left + right
