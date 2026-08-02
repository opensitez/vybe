# vybe-test: python/stdlib_compile_extended/contextlib_asynccontextmanager
# origin: languages/python/tests/python/test_stdlib_compile_extended.rs
# vybe-test-mode: compile

from contextlib import asynccontextmanager
@asynccontextmanager
async def cm():
 yield 1
