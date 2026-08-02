# vybe-test: python/context_manager_spec/ctx_tuple_target_compile
# origin: languages/python/tests/python/test_context_manager_spec.rs
# vybe-test-mode: compile

class Ctx:
    def __enter__(self):
        return (1, 2)
    def __exit__(self, exc_type, exc, tb):
        return False
with Ctx() as (a, b):
    print(a, b)
