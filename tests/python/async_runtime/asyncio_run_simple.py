# vybe-test: python/async_runtime/asyncio_run_simple
# origin: languages/python/tests/python/test_async_runtime.rs
# vybe-test-mode: compile

import asyncio
async def main():
 return 1
asyncio.run(main())
