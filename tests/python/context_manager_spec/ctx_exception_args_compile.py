# vybe-test: python/context_manager_spec/ctx_exception_args_compile
# origin: languages/python/tests/python/test_context_manager_spec.rs
# vybe-test-mode: compile

class Ctx:
    def __enter__(self):
        return self
    def __exit__(self, exc_type, exc, tb):
        print(exc_type, exc, tb)
        return False
