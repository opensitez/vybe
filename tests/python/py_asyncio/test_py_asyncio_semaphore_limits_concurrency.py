# vybe-test: python/py_asyncio/test_py_asyncio_semaphore_limits_concurrency
# origin: languages/python/tests/python/test_py_asyncio.rs

import asyncio

active = [0]
max_active = [0]
sem = asyncio.Semaphore(2)

async def task():
    async with sem:
        active[0] += 1
        max_active[0] = max(max_active[0], active[0])
        await asyncio.sleep(0)
        active[0] -= 1

async def main():
    await asyncio.gather(*[task() for _ in range(5)])
    print(max_active[0] <= 2)

asyncio.run(main())
