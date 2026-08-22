# vybe-test: python/function_signatures_spec/async_generator_with_await_compile
# origin: languages/python/tests/python/test_function_signatures_spec.rs

async def gen():
    value = await fetch()
    yield value
