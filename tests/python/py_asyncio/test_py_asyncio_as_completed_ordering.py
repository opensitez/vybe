# vybe-test: python/py_asyncio/test_py_asyncio_as_completed_ordering
# origin: languages/python/tests/python/test_py_asyncio.rs

import asyncio

async def delayed(val, delay):
    await asyncio.sleep(delay)
    return val

async def main():
    coros = [delayed("slow", 0.03), delayed("fast", 0.01), delayed("mid", 0.02)]
    results = []
    for coro in asyncio.as_completed(coros):
        results.append(await coro)
    print(results)

asyncio.run(main())
