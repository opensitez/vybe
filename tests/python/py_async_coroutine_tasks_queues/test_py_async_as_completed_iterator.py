# vybe-test: python/py_async_coroutine_tasks_queues/test_py_async_as_completed_iterator
# origin: languages/python/tests/python/test_py_async_coroutine_tasks_queues.rs

import asyncio

async def delay_return(val, delay):
    await asyncio.sleep(delay)
    return val

async def main():
    tasks = [delay_return(3, 0.003), delay_return(1, 0.001), delay_return(2, 0.002)]
    results = []
    for fut in asyncio.as_completed(tasks):
        res = await fut
        results.append(res)
    print(results)  # should arrive in order of completion: [1, 2, 3]

asyncio.run(main())
