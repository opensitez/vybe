# vybe-test: python/context_manager_spec/ctx_enter_value_compile
# origin: languages/python/tests/python/test_context_manager_spec.rs

class Ctx:
    def __enter__(self):
        return 42
    def __exit__(self, exc_type, exc, tb):
        return False
with Ctx() as value:
    print(value)
