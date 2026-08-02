# vybe-test: python/for_while_extended/while_async
# origin: languages/python/tests/python/test_for_while_extended.rs
# vybe-test-mode: compile

async def f():
 while True:
  await asyncio.sleep(0)
  break
