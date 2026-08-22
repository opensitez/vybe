# vybe-test: python/function_signatures_extended/inspect_bound_arguments
# origin: languages/python/tests/python/test_function_signatures_extended.rs

import inspect
def f(a, b=1): pass
inspect.signature(f).bind(1)
