# vybe-test: python/closure_truthiness_spec/closure_default_bind_compile
# origin: languages/python/tests/python/test_closure_truthiness_spec.rs
# vybe-test-mode: compile

funcs = [lambda x=i: x for i in range(3)]
