# vybe-test: python/unpacking_extended/unpack_async_for
# origin: languages/python/tests/python/test_unpacking_extended.rs

async def f():
 async for x, in async_items():
  pass
