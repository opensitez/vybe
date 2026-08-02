# vybe-test: python/function_signatures_spec/async_comprehension_if_compile
# origin: languages/python/tests/python/test_function_signatures_spec.rs
# vybe-test-mode: compile

result = [x async for x in aiter if x > 0]
