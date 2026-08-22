# vybe-test: python/context_manager_spec/ctx_async_generator_style_compile
# origin: languages/python/tests/python/test_context_manager_spec.rs

from contextlib import asynccontextmanager
@asynccontextmanager
async def acm():
    yield 1
