# vybe-test: python/python_asyncio_tasks_threads/test_asyncio_as_completed_yields_futures
# origin: languages/python/tests/python/test_python_asyncio_tasks_threads.rs

import asyncio

async def slow():
    await asyncio.sleep(0.02)
    return "slow"

async def fast():
    await asyncio.sleep(0.005)
    return "fast"

async def fn():
    results = []
    for coro in asyncio.as_completed([slow(), fast()]):
        val = await coro
        results.append(val)
    print(results)

asyncio.run(fn())
