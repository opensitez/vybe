# vybe-test: python/py_asyncio/test_py_asyncio_queue_producer_consumer
# origin: languages/python/tests/python/test_py_asyncio.rs

import asyncio

consumed = []

async def producer(queue):
    for i in range(3):
        await queue.put(i)
    await queue.put(None)  # sentinel

async def consumer(queue):
    while True:
        item = await queue.get()
        if item is None:
            break
        consumed.append(item)

async def main():
    q = asyncio.Queue()
    await asyncio.gather(producer(q), consumer(q))
    print(consumed)

asyncio.run(main())
