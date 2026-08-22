# vybe-test: python/context_manager_spec/ctx_return_none_compile
# origin: languages/python/tests/python/test_context_manager_spec.rs

class Ctx:
    def __enter__(self):
        return None
    def __exit__(self, exc_type, exc, tb):
        return False
with Ctx() as value:
    pass
