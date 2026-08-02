# vybe-test: python/walrus_extended/walrus_async_comp
# origin: languages/python/tests/python/test_walrus_extended.rs
# vybe-test-mode: compile

async def f():
 return [(x := i) async for i in async_range(2)]
