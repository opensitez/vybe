# vybe-test: python/context_manager_spec/ctx_in_function_compile
# origin: languages/python/tests/python/test_context_manager_spec.rs
# vybe-test-mode: compile

def read(name):
    with open(name) as f:
        return f.read()
