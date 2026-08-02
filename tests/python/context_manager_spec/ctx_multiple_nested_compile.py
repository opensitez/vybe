# vybe-test: python/context_manager_spec/ctx_multiple_nested_compile
# origin: languages/python/tests/python/test_context_manager_spec.rs
# vybe-test-mode: compile

with open('a') as f:
    with open('b') as g:
        with open('c') as h:
            pass
