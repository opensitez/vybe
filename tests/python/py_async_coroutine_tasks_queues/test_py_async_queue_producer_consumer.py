# vybe-test: python/py_async_coroutine_tasks_queues/test_py_async_queue_producer_consumer
# origin: languages/python/tests/python/test_py_async_coroutine_tasks_queues.rs

import asyncio

async def producer(queue):
    for i in range(3):
        await queue.put(i)

async def consumer(queue, out):
    while True:
        item = await queue.get()
        out.append(item)
        queue.task_done()
        if len(out) == 3:
            break

async def main():
    q = asyncio.Queue()
    out = []
    prod_task = asyncio.create_task(producer(q))
    cons_task = asyncio.create_task(consumer(q, out))
    await prod_task
    await cons_task
    print(out)

asyncio.run(main())
