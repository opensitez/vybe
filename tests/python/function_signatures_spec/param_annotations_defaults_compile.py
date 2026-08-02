# vybe-test: python/function_signatures_spec/param_annotations_defaults_compile
# origin: languages/python/tests/python/test_function_signatures_spec.rs
# vybe-test-mode: compile

def f(a: int = 1, b: str = 'x') -> str:
    return b
