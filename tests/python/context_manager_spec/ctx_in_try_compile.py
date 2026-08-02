# vybe-test: python/context_manager_spec/ctx_in_try_compile
# origin: languages/python/tests/python/test_context_manager_spec.rs
# vybe-test-mode: compile

try:
    with open('x') as f:
        pass
finally:
    cleanup()
