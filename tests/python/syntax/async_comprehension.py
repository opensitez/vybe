# vybe-test: python/syntax/async_comprehension
# origin: languages/python/tests/python/test_syntax.rs
# vybe-test-mode: compile

result = [x async for x in aiter]
