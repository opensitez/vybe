# vybe-test: python/context_manager_spec/ctx_in_loop_compile
# origin: languages/python/tests/python/test_context_manager_spec.rs
# vybe-test-mode: compile

for name in ['a', 'b']:
    with open(name) as f:
        print(name)
