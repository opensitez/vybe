# vybe-test: python/generator_iterator_protocol_spec/async_generator_yield_from_like_compile
# origin: languages/python/tests/python/test_generator_iterator_protocol_spec.rs
# vybe-test-mode: compile

async def agen():
    for x in [1, 2]:
        yield x
