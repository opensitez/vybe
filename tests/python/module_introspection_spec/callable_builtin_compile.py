# vybe-test: python/module_introspection_spec/callable_builtin_compile
# origin: languages/python/tests/python/test_module_introspection_spec.rs
# vybe-test-mode: compile

def f():
    pass
ok = callable(f)
