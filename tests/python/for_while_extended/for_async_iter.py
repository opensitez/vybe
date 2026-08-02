# vybe-test: python/for_while_extended/for_async_iter
# origin: languages/python/tests/python/test_for_while_extended.rs
# vybe-test-mode: compile

async def f():
 async for x in agen():
  pass
