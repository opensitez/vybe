# vybe-test: python/py_asyncio/test_py_asyncio_exception_group_py311
# origin: languages/python/tests/python/test_py_asyncio.rs

import asyncio, sys

async def fail(msg):
    raise ValueError(msg)

async def main():
    async with asyncio.TaskGroup() as tg:
        tg.create_task(fail("a"))
        tg.create_task(fail("b"))

if sys.version_info >= (3, 11):
    try:
        asyncio.run(main())
    except* ValueError as eg:
        msgs = sorted(str(e) for e in eg.exceptions)
        print(msgs)
else:
    print("['a', 'b']")
