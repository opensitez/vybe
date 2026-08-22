# vybe-test: python/function_signatures_spec/bare_star_kwonly_compile
# origin: languages/python/tests/python/test_function_signatures_spec.rs

def f(*, flag=False):
    return flag
