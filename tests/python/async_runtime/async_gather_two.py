# vybe-test: python/async_runtime/async_gather_two
# origin: languages/python/tests/python/test_async_runtime.rs
# vybe-test-mode: compile

import asyncio
async def f():
 return 1
async def g():
 return 2
asyncio.run(asyncio.gather(f(), g()))
