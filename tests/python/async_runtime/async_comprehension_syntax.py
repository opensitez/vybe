# vybe-test: python/async_runtime/async_comprehension_syntax
# origin: languages/python/tests/python/test_async_runtime.rs
# vybe-test-mode: compile

async def f():
 return [i async for i in aiter([])]
