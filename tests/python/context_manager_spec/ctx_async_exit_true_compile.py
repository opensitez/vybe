# vybe-test: python/context_manager_spec/ctx_async_exit_true_compile
# origin: languages/python/tests/python/test_context_manager_spec.rs

class ACtx:
    async def __aenter__(self):
        return self
    async def __aexit__(self, exc_type, exc, tb):
        return True
