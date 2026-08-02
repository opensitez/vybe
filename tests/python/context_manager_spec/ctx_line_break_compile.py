# vybe-test: python/context_manager_spec/ctx_line_break_compile
# origin: languages/python/tests/python/test_context_manager_spec.rs
# vybe-test-mode: compile

with open('a') as f,
     open('b') as g:
    pass
