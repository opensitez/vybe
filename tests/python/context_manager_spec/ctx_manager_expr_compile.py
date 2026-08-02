# vybe-test: python/context_manager_spec/ctx_manager_expr_compile
# origin: languages/python/tests/python/test_context_manager_spec.rs
# vybe-test-mode: compile

factory = open
with factory('x') as f:
    pass
