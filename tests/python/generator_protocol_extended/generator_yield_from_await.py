# vybe-test: python/generator_protocol_extended/generator_yield_from_await
# origin: languages/python/tests/python/test_generator_protocol_extended.rs
# vybe-test-mode: compile

async def ag():
 yield from async_iter()
