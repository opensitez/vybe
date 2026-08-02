# vybe-test: python/syntax/async_for
# origin: languages/python/tests/python/test_syntax.rs
# vybe-test-mode: compile

async def main():
    async for item in aiter:
        print(item)
