# vybe-test: python/async_runtime/async_task_group
# origin: languages/python/tests/python/test_async_runtime.rs

import asyncio
async def main():
 async with asyncio.TaskGroup() as tg:
  pass
asyncio.run(main())
