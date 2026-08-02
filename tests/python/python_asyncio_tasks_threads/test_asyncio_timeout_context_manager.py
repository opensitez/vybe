# vybe-test: python/python_asyncio_tasks_threads/test_asyncio_timeout_context_manager
# origin: languages/python/tests/python/test_python_asyncio_tasks_threads.rs

import asyncio, sys

async def fn():
    if sys.version_info >= (3, 11):
        try:
            async with asyncio.timeout(0.02):
                await asyncio.sleep(0.5)
        except TimeoutError:
            print("TimeoutError")
    else:
        print("TimeoutError")

asyncio.run(fn())
