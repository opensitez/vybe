# vybe-test: python/function_signatures_spec/assert_in_except_compile
# origin: languages/python/tests/python/test_function_signatures_spec.rs

try:
    risky()
except Exception as exc:
    assert exc is not None
