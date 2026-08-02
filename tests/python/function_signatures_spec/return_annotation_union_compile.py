# vybe-test: python/function_signatures_spec/return_annotation_union_compile
# origin: languages/python/tests/python/test_function_signatures_spec.rs
# vybe-test-mode: compile

def f(x: int) -> int | None:
    return x
