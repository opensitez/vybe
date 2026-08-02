# vybe-test: python/module_introspection_spec/locals_builtin_compile
# origin: languages/python/tests/python/test_module_introspection_spec.rs
# vybe-test-mode: compile

def f():
    x = 1
    return locals()
