# vybe-test: python/function_signatures_spec/decorator_with_return_annotation_compile
# origin: languages/python/tests/python/test_function_signatures_spec.rs
# vybe-test-mode: compile

@cache
def f(x: int) -> int:
    return x
