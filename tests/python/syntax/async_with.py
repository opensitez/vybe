# vybe-test: python/syntax/async_with
# origin: languages/python/tests/python/test_syntax.rs
# vybe-test-mode: compile

async def ctx():
    async with resource() as r:
        pass
