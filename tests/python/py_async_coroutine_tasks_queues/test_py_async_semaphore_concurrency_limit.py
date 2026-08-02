# vybe-test: python/py_async_coroutine_tasks_queues/test_py_async_semaphore_concurrency_limit
# origin: languages/python/tests/python/test_py_async_coroutine_tasks_queues.rs

import asyncio

active = 0
max_active = 0

async def worker(sem):
    global active, max_active
    async with sem:
        active += 1
        if active > max_active:
            max_active = active
        await asyncio.sleep(0.001)
        active -= 1

async def main():
    sem = asyncio.Semaphore(2)
    tasks = [asyncio.create_task(worker(sem)) for _ in range(5)]
    await asyncio.gather(*tasks)
    print(max_active <= 2)

asyncio.run(main())
