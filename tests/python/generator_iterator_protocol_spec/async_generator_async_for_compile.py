# vybe-test: python/generator_iterator_protocol_spec/async_generator_async_for_compile
# origin: languages/python/tests/python/test_generator_iterator_protocol_spec.rs
# vybe-test-mode: compile

async def main():
    async for value in agen():
        print(value)
