# vybe-test: python/async_runtime/async_with_contextlib
# origin: languages/python/tests/python/test_async_runtime.rs

from contextlib import asynccontextmanager
@asynccontextmanager
async def cm():
 yield 1
