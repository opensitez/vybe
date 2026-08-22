# vybe-test: python/function_signatures_spec/def_with_annotations_and_unpack_compile
# origin: languages/python/tests/python/test_function_signatures_spec.rs

def f(a: int, *args: str, **kwargs: float) -> tuple[int, tuple[str], dict[str, float]]:
    return a, args, kwargs
