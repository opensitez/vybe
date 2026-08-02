# vybe-test: python/function_signatures_extended/type_params_pep695
# origin: languages/python/tests/python/test_function_signatures_extended.rs
# vybe-test-mode: compile

def f[T](x: T) -> T: return x
