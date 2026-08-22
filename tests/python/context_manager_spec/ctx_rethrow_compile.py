# vybe-test: python/context_manager_spec/ctx_rethrow_compile
# origin: languages/python/tests/python/test_context_manager_spec.rs
def risky(*_a, **_k):
    return None

class Ctx:
    def __enter__(self):
        return self
    def __exit__(self, exc_type, exc, tb):
        return False
with Ctx():
    risky()
