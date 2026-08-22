# vybe-test: python/async_runtime/async_wait_for
# origin: languages/python/tests/python/test_async_runtime.rs

import asyncio
async def f():
 return 1
asyncio.run(asyncio.wait_for(f(), timeout=1))
