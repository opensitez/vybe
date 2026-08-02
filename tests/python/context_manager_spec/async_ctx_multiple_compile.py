# vybe-test: python/context_manager_spec/async_ctx_multiple_compile
# origin: languages/python/tests/python/test_context_manager_spec.rs
# vybe-test-mode: compile

async def main():
    async with first() as a, second() as b:
        print(a, b)
