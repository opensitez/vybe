# vybe-test: python/context_manager_spec/ctx_no_as_compile
# origin: languages/python/tests/python/test_context_manager_spec.rs
# vybe-test-mode: compile

class Lock:
    def __enter__(self):
        return self
    def __exit__(self, exc_type, exc, tb):
        return False
with Lock():
    pass
