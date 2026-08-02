# vybe-test: python/context_managers_extended/contextlib_asynccontextmanager
# origin: languages/python/tests/python/test_context_managers_extended.rs
# vybe-test-mode: compile

from contextlib import asynccontextmanager
@asynccontextmanager
async def cm():
 yield 1
