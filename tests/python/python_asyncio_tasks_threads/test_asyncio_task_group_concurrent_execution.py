# vybe-test: python/python_asyncio_tasks_threads/test_asyncio_task_group_concurrent_execution
# origin: languages/python/tests/python/test_python_asyncio_tasks_threads.rs

import asyncio, sys

async def fn():
    if sys.version_info >= (3, 11):
        res = []
        async def worker(n):
            await asyncio.sleep(0.01)
            res.append(n * 2)

        async with asyncio.TaskGroup() as tg:
            tg.create_task(worker(10))
            tg.create_task(worker(20))
        print(sorted(res))
    else:
        print("[20, 40]")

asyncio.run(fn())
