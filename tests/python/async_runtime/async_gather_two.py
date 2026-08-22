# vybe-test: python/async_runtime/async_gather_two
# origin: languages/python/tests/python/test_async_runtime.rs
# Coroutines must be created and awaited INSIDE the running loop —
# `asyncio.run(asyncio.gather(...))` builds them outside it and leaves
# them never-awaited.
import asyncio
async def f():
 return 1
async def g():
 return 2
async def main():
 return await asyncio.gather(f(), g())
asyncio.run(main())
