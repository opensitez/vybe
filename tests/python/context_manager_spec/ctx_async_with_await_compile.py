# vybe-test: python/context_manager_spec/ctx_async_with_await_compile
# origin: languages/python/tests/python/test_context_manager_spec.rs
# vybe-test-mode: compile

async def main():
    async with await make_ctx() as ctx:
        print(ctx)
