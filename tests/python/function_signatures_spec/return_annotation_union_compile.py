# vybe-test: python/function_signatures_spec/return_annotation_union_compile
# origin: languages/python/tests/python/test_function_signatures_spec.rs

def f(x: int) -> int | None:
    return x
