# vybe-test: python/syntax/async_def
# origin: languages/python/tests/python/test_syntax.rs

async def fetch():
    data = await get_data()
    return data
