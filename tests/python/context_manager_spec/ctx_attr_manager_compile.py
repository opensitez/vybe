# vybe-test: python/context_manager_spec/ctx_attr_manager_compile
# origin: languages/python/tests/python/test_context_manager_spec.rs
# vybe-test-mode: compile

obj = resource_holder
with obj.manager() as r:
    pass
