# vybe-test: python/py_asyncio/test_py_async_generator_async_for
# origin: languages/python/tests/python/test_py_asyncio.rs

import asyncio

async def async_range(n):
    for i in range(n):
        await asyncio.sleep(0)
        yield i

async def main():
    results = []
    async for val in async_range(4):
        results.append(val)
    print(results)

asyncio.run(main())
