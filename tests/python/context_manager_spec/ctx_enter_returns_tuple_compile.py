# vybe-test: python/context_manager_spec/ctx_enter_returns_tuple_compile
# origin: languages/python/tests/python/test_context_manager_spec.rs
# vybe-test-mode: compile

class Ctx:
    def __enter__(self):
        return (1, 2, 3)
    def __exit__(self, exc_type, exc, tb):
        return False
