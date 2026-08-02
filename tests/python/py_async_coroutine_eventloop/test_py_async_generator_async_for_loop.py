# vybe-test: python/py_async_coroutine_eventloop/test_py_async_generator_async_for_loop
# origin: languages/python/tests/python/test_py_async_coroutine_eventloop.rs

import asyncio

async def async_numbers(n):
    for i in range(n):
        await asyncio.sleep(0.001)
        yield i

async def main():
    out = []
    async for num in async_numbers(3):
        out.append(num)
    print(out)

asyncio.run(main())
